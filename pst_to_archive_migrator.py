import win32com.client
import pythoncom
import logging
import time
from datetime import datetime
import json
import os
import tkinter as tk
from tkinter import messagebox

class EmailMigrator:
    def __init__(self):
        self.setup_logging()
        self.migration_report = {
            'start_time': None,
            'end_time': None,
            'destination_type': 'N/A',
            'destination_path': 'N/A',
            'total_attempted': 0,
            'total_successful': 0,
            'total_failed': 0,
            'aggregate_validation_passed': False,
            'pst_migrations_details': [],
            'failed_items_overall': [],
            'folder_summary_overall': {}
        }

    def setup_logging(self):
        """Setup comprehensive logging"""
        log_dir = "migration_logs"
        os.makedirs(log_dir, exist_ok=True)
        log_filename = os.path.join(log_dir, f'pst_migration_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')

        logging.basicConfig(
            filename=log_filename,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        formatter = logging.Formatter('%(levelname)s: %(message)s')
        console.setFormatter(formatter)
        logging.getLogger().addHandler(console)

        file_handler = logging.FileHandler(log_filename)
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(file_handler)

        logging.info(f"Logging to: {log_filename}")

    def get_item_signature(self, item):
        """Create a unique signature for reporting/logging"""
        try:
            return {
                'subject': getattr(item, 'Subject', 'N/A'),
                'sent_on': str(getattr(item, 'SentOn', None)) if getattr(item, 'SentOn', None) else None,
                'sender': getattr(item, 'SenderName', 'N/A'),
                'entry_id': getattr(item, 'EntryID', 'N/A'),
                'size': getattr(item, 'Size', 'N/A')
            }
        except Exception as e:
            return {'subject': 'Unknown', 'error': f'Could not get item properties: {e}'}

    def get_folder_item_count(self, folder_obj):
        """Safely get the count of items in an Outlook folder."""
        try:
            return folder_obj.Items.Count
        except Exception as e:
            logging.error(f"Error getting item count for folder '{getattr(folder_obj, 'FolderPath', 'N/A')}': {e}", exc_info=True)
            return -1

    def process_folder(self, source_folder, target_folder, current_pst_report, folder_path=""):
        """
        Process all items in a folder and move them to the specified target_folder.
        Updates the current_pst_report with counts for the current PST.
        """

        current_folder_path = f"{folder_path}/{source_folder.Name}" if folder_path else source_folder.Name
        logging.info(f"Processing folder: {current_folder_path}")

        try:
            items_list = list(source_folder.Items)
            item_count = len(items_list)
            logging.info(f"Found {item_count} items in '{current_folder_path}' to process.")

            for item in reversed(items_list):
                try:
                    if item.Class == 43:
                        current_pst_report['total_attempted_current_pst'] += 1

                        original_signature = self.get_item_signature(item)
                        logging.info(f"Attempting to move item: {original_signature['subject'][:50]}...")

                        try:
                            item.Move(target_folder)

                            current_pst_report['total_successful_current_pst'] += 1
                            logging.info(f"Successfully moved: {original_signature['subject'][:50]}...")

                        except Exception as move_error:
                            current_pst_report['total_failed_current_pst'] += 1
                            current_pst_report['failed_items_current_pst'].append({
                                'folder': current_folder_path,
                                'subject': original_signature['subject'],
                                'error': str(move_error),
                                'original_signature': original_signature
                            })
                            logging.error(f"Move failed for '{original_signature['subject']}': {move_error}", exc_info=True)

                except Exception as item_error:
                    current_pst_report['total_failed_current_pst'] += 1
                    logging.error(f"Error processing item in '{current_folder_path}': {item_error}", exc_info=True)
                    continue

        except Exception as folder_items_error:
            logging.error(f"Error accessing items in folder '{current_folder_path}': {folder_items_error}", exc_info=True)
            current_pst_report['total_failed_current_pst'] += item_count if 'item_count' in locals() else 0

        current_pst_report['folder_summary_current_pst'][current_folder_path] = {
            'successful_moves_from_this_folder': current_pst_report['total_successful_current_pst'] - sum(f['successful_moves_from_this_folder'] for f in current_pst_report['folder_summary_current_pst'].values()),
            'failed_moves_from_this_folder': current_pst_report['total_failed_current_pst'] - sum(f['failed_moves_from_this_folder'] for f in current_pst_report['folder_summary_current_pst'].values()),
            'total_items_in_source_folder': item_count if 'item_count' in locals() else 0
        }

        for subfolder in source_folder.Folders:
            try:
                self.process_folder(subfolder, target_folder, current_pst_report, current_folder_path)
            except Exception as subfolder_error:
                logging.error(f"Error processing subfolder '{subfolder.Name}': {subfolder_error}", exc_info=True)
                current_pst_report['total_failed_current_pst'] += 1

    def generate_report(self):
        """Generate a comprehensive migration report"""
        report_dir = "migration_reports"
        os.makedirs(report_dir, exist_ok=True)
        report_filename = os.path.join(report_dir, f'migration_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')

        self.migration_report['end_time'] = datetime.now().isoformat()
        if self.migration_report['start_time']:
            self.migration_report['duration_seconds'] = (
                datetime.fromisoformat(self.migration_report['end_time']) -
                datetime.fromisoformat(self.migration_report['start_time'])
            ).total_seconds()
        else:
            self.migration_report['duration_seconds'] = 0

        self.migration_report['success_rate'] = (
            self.migration_report['total_successful'] / self.migration_report['total_attempted'] * 100
        ) if self.migration_report['total_attempted'] > 0 else 0

        with open(report_filename, 'w') as f:
            json.dump(self.migration_report, f, indent=2, default=str)

        logging.info(f"Migration report saved to: {report_filename}")

        print(f"\n=== OVERALL MIGRATION SUMMARY ===")
        print(f"Destination Type: {self.migration_report['destination_type']}")
        print(f"Destination Path: {self.migration_report['destination_path']}")
        print(f"Start Time: {self.migration_report['start_time']}")
        print(f"End Time: {self.migration_report['end_time']}")
        print(f"Duration: {self.migration_report['duration_seconds']:.2f} seconds")
        print(f"Grand Total PST Items Found (Attempted to Move): {self.migration_report['total_attempted']}")
        print(f"Grand Total Successful .Move() Operations: {self.migration_report['total_successful']}")
        print(f"Grand Total Failed .Move() Operations: {self.migration_report['total_failed']}")
        print(f"Overall Success Rate (of .Move() calls): {self.migration_report['success_rate']:.2f}%")
        print(f"Aggregate Validation Passed: {self.migration_report['aggregate_validation_passed']}")

        if self.migration_report['failed_items_overall']:
            print(f"\n=== OVERALL FAILED ITEMS ({len(self.migration_report['failed_items_overall'])}) ===")
            print("Showing first 10 overall failures (check log file for all details):")
            for i, failed_item in enumerate(self.migration_report['failed_items_overall'][:10], 1):
                print(f"{i}. PST: {failed_item.get('pst_display_name', 'N/A')}, Subject: {failed_item['subject'][:70]}...")
                print(f"   From Folder: {failed_item['folder']}")
                print(f"   Error: {failed_item['error']}")
                print("-" * 30)

        print(f"\n=== INDIVIDUAL PST MIGRATION DETAILS ===")
        if not self.migration_report['pst_migrations_details']:
            print("No PSTs were processed.")
        else:
            for pst_detail in self.migration_report['pst_migrations_details']:
                print(f"\n--- PST: '{pst_detail['pst_display_name']}' (Path: {pst_detail['pst_file_path']}) ---")
                print(f"  Items Attempted: {pst_detail['total_attempted_current_pst']}")
                print(f"  Successful Moves: {pst_detail['total_successful_current_pst']}")
                print(f"  Failed Moves: {pst_detail['total_failed_current_pst']}")
                pst_success_rate = (
                    pst_detail['total_successful_current_pst'] / pst_detail['total_attempted_current_pst'] * 100
                ) if pst_detail['total_attempted_current_pst'] > 0 else 0
                print(f"  Success Rate: {pst_success_rate:.2f}%")

                if pst_detail['failed_items_current_pst']:
                    print(f"  Failed items for this PST ({len(pst_detail['failed_items_current_pst'])}):")
                    for i, failed_item in enumerate(pst_detail['failed_items_current_pst'][:5], 1):
                        print(f"    {i}. Subject: {failed_item['subject'][:70]}... - {failed_item['error']}")
                else:
                    print("  No failed items for this PST.")
                print("-" * 50)


    def select_pst_store(self, namespace):
        """Identifies and returns a list of all open PST stores."""
        pst_stores = []
        for store in namespace.Stores:
            if hasattr(store, 'FilePath') and store.FilePath and store.FilePath.lower().endswith('.pst'):
                pst_stores.append(store)

        if not pst_stores:
            logging.error("No PST files found currently open in Outlook.")
            messagebox.showerror("No PSTs Open", "No PST files are currently open in Outlook. Please open the desired PST file(s) in Outlook first, then run this script.")
            return None

        print("\n--- Detected PST Files Open in Outlook ---")
        for i, store in enumerate(pst_stores):
            print(f"{i + 1}. {store.DisplayName} (Path: {getattr(store, 'FilePath', 'N/A')})")

        return pst_stores

    def select_destination_store(self, namespace):
        """
        Prompts the user to select any open mailbox/store as the destination.
        Returns the root folder of the selected store.
        """
        all_stores = list(namespace.Stores)
        if not all_stores:
            logging.error("No mailboxes or stores found in Outlook.")
            messagebox.showerror("No Mailboxes Found", "No mailboxes or stores are currently open in Outlook. Please ensure Outlook is running with at least one mailbox open.")
            return None, None, None

        print("\n--- Select Migration Destination ---")
        display_list = []
        for i, store in enumerate(all_stores):
            try:
                root_folder = store.GetRootFolder()
                display_list.append((store, root_folder))
                print(f"{len(display_list)}. {store.DisplayName} (Path: {root_folder.FolderPath})")
            except Exception as e:
                logging.warning(f"Could not access root folder for store '{store.DisplayName}', skipping: {e}")

        if not display_list:
            logging.error("Could not find any selectable destinations.")
            messagebox.showerror("No Destinations Found", "Could not find any folders to migrate to. Check Outlook's configuration.")
            return None, None, None

        while True:
            try:
                choice = input("Enter the number of the destination mailbox: ")
                selected_index = int(choice) - 1

                if 0 <= selected_index < len(display_list):
                    selected_store, target_root_folder = display_list[selected_index]
                    logging.info(f"Target destination selected: {target_root_folder.FolderPath}")
                    return target_root_folder, selected_store.DisplayName, target_root_folder.FolderPath
                else:
                    print("Invalid choice. Please enter a valid number from the list.")
            except ValueError:
                print("Invalid input. Please enter a number.")
            except Exception as e:
                print(f"Error during destination selection: {e}. Please try again.")
                logging.error(f"Unexpected error during destination selection: {e}", exc_info=True)
                return None, None, None

    def run_migration(self):
        """Main migration function with comprehensive validation for multiple PSTs"""
        logging.info("Starting PST to Destination migration.")
        self.migration_report['start_time'] = datetime.now().isoformat()

        pythoncom.CoInitialize()
        outlook = None
        all_pst_stores = []
        target_folder = None
        target_display_name = None
        target_path = None

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            # --- Step 1: Detect all open PSTs ---
            all_pst_stores = self.select_pst_store(namespace)
            if not all_pst_stores:
                logging.error("No PST stores found to process. Exiting migration.")
                return False

            logging.info(f"Detected {len(all_pst_stores)} PSTs for migration.")
            source_pst_file_paths = [getattr(store, 'FilePath', 'N/A') for store in all_pst_stores]

            # --- Step 2: User selects the destination type ---
            target_folder, target_display_name, target_path = self.select_destination_store(namespace)

            if not target_folder:
                logging.error("Target destination selection failed. Exiting migration.")
                return False

            self.migration_report['destination_type'] = "User Selected Store"
            self.migration_report['destination_path'] = target_path

            # --- Step 3: Get initial count of items in the target folder ---
            initial_target_item_count = self.get_folder_item_count(target_folder)
            if initial_target_item_count == -1:
                logging.error("Failed to get initial item count for target destination. Cannot perform aggregate validation.")
                messagebox.showerror("Validation Error", "Failed to get initial item count. Check logs.")
                return False
            logging.info(f"Initial item count in target destination ('{target_path}'): {initial_target_item_count}")
            print(f"\nInitial item count in target destination: {initial_target_item_count}")

            print("\n" + "=" * 50)
            print(f"CONFIRMATION: You are about to MOVE ALL emails from {len(all_pst_stores)} selected PST(s)")
            print(f"into the '{self.migration_report['destination_type']}' of '{target_display_name}' ('{target_path}').")
            print("This action will flatten the folder structures of ALL PSTs into the single target folder.")
            print("THIS ACTION CANNOT BE UNDONE. ENSURE YOU HAVE BACKUPS OF ALL YOUR PST FILES.")
            print("=" * 50)

            confirm = input("Type 'CONFIRM' to proceed with the migration: ")
            if confirm.strip().upper() != 'CONFIRM':
                logging.warning("Migration cancelled by user.")
                print("Operation cancelled.")
                return False

            # --- Step 4: Process each PST ---
            for i, pst_store_obj in enumerate(all_pst_stores):
                pst_store_display_name = pst_store_obj.DisplayName
                pst_file_path = getattr(pst_store_obj, 'FilePath', 'N/A')

                logging.info(f"\n--- Processing PST {i+1}/{len(all_pst_stores)}: '{pst_store_display_name}' (Path: {pst_file_path}) ---")
                print(f"\n--- Processing PST {i+1}/{len(all_pst_stores)}: '{pst_store_display_name}' ---")

                current_pst_report = {
                    'pst_display_name': pst_store_display_name,
                    'pst_file_path': pst_file_path,
                    'total_attempted_current_pst': 0,
                    'total_successful_current_pst': 0,
                    'total_failed_current_pst': 0,
                    'failed_items_current_pst': [],
                    'folder_summary_current_pst': {}
                }

                self.process_folder(
                    pst_store_obj.GetRootFolder(),
                    target_folder,
                    current_pst_report
                )

                self.migration_report['pst_migrations_details'].append(current_pst_report)
                self.migration_report['total_attempted'] += current_pst_report['total_attempted_current_pst']
                self.migration_report['total_successful'] += current_pst_report['total_successful_current_pst']
                self.migration_report['total_failed'] += current_pst_report['total_failed_current_pst']
                self.migration_report['failed_items_overall'].extend(
                    [{**item, 'pst_display_name': pst_store_display_name} for item in current_pst_report['failed_items_current_pst']]
                )
                self.migration_report['folder_summary_overall'].update(current_pst_report['folder_summary_current_pst'])

            logging.info(f"\nFinished moving items from all {len(all_pst_stores)} PSTs.")
            print(f"\nFinished moving items from all {len(all_pst_stores)} PSTs.")

            # --- Step 5: Post-migration wait for synchronization ---
            logging.info(f"Waiting for 10 seconds to allow target destination to synchronize...")
            print("Waiting for 60 seconds for target to synchronize...")
            time.sleep(5)
            logging.info("Wait complete.")

            # --- Step 6: Post-migration count and aggregate validation ---
            final_target_item_count = self.get_folder_item_count(target_folder)
            if final_target_item_count == -1:
                logging.error("Failed to get final item count for target destination. Aggregate validation cannot be completed.")
                messagebox.showerror("Validation Error", "Failed to get final item count. Check logs.")
                self.migration_report['aggregate_validation_passed'] = False
            else:
                logging.info(f"Final item count in target destination ('{target_path}'): {final_target_item_count}")
                print(f"Final item count in target destination: {final_target_item_count}")

                delta_items = final_target_item_count - initial_target_item_count
                logging.info(f"Delta items in target: {delta_items}. Total successful moves across all PSTs: {self.migration_report['total_successful']}")

                if (final_target_item_count > initial_target_item_count and
                    delta_items >= self.migration_report['total_successful']):
                    self.migration_report['aggregate_validation_passed'] = True
                    logging.info("Aggregate count validation PASSED: Target increased by expected amount.")
                else:
                    self.migration_report['aggregate_validation_passed'] = False
                    logging.warning(
                        f"Aggregate count validation FAILED. "
                        f"Expected increase >= {self.migration_report['total_successful']}, but got {delta_items}. "
                        f"(Initial: {initial_target_item_count}, Final: {final_target_item_count})"
                    )
                    messagebox.showwarning(
                        "Aggregate Validation Failed",
                        f"Count validation failed.\n"
                        f"Expected total increase from all PSTs: >= {self.migration_report['total_successful']}\n"
                        f"Actual total increase in destination: {delta_items}\n"
                        f"Please check Outlook and logs for discrepancies."
                    )

            self.generate_report()

            overall_success = (self.migration_report['total_failed'] == 0 and self.migration_report['aggregate_validation_passed'])

            if overall_success:
                logging.info("Migration completed successfully with 0 errors and passed aggregate validation!")
            else:
                logging.warning(f"Migration completed with {self.migration_report['total_failed']} errors. Aggregate validation: {self.migration_report['aggregate_validation_passed']}. Please check the log and report files.")

            return overall_success

        except Exception as e:
            logging.error(f"Critical error during migration: {e}", exc_info=True)
            messagebox.showerror("Migration Error", f"A critical error occurred during migration. Check logs for details: {e}")
            return False
        finally:
            if outlook:
                del outlook
            pythoncom.CoUninitialize()

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("   PST to Destination Email Migration Tool (All Open PSTs)   ")
    print("=" * 60)
    print("This tool requires you to **manually open the desired PST file(s) in Outlook first**.")
    print("It will then list all these open PSTs and process them sequentially,")
    print("moving emails from all their folders to a single selected destination.")
    print("\n" + "=" * 60)

    migrator = EmailMigrator()
    success = migrator.run_migration()

    if success:
        print("\n✅ Migration completed successfully with no errors and passed aggregate validation!")
    else:
        print("\n❌ Migration completed with errors or failed aggregate validation. Check 'migration_logs' and 'migration_reports' folders for details.")
