import pandas as pd
from geopy.distance import geodesic
from fuzzywuzzy import process, fuzz
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, PhotoImage
import threading
import time
from openpyxl.styles import PatternFill, Font

class ResortMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Resort Matcher")
        self.root.geometry("1024x900")

        logo_image = PhotoImage(file="logo.png")
        logo_label = tk.Label(self.root, image=logo_image)
        logo_label.image = logo_image
        logo_label.pack(pady=10)

        ttk.Button(self.root, text="Select Excel File", command=self.load_excel).pack(pady=10)
        self.save_results_button = ttk.Button(self.root, text="Save Results", command=self.save_results)
        self.save_results_button.pack_forget()

        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress.pack(pady=20)

        self.time_remaining_label = ttk.Label(self.root, text="Estimated Time Remaining: N/A")
        self.time_remaining_label.pack(pady=10)

        self.processed_rows_label = ttk.Label(self.root, text="Rows Processed: 0 / 0")
        self.processed_rows_label.pack(pady=10)

        self.cancel_process = False
        self.cancel_button = ttk.Button(self.root, text="Cancel", command=self.cancel_operation)
        self.cancel_button.pack(pady=10)

        # Setup for the Treeview
        self.tree = ttk.Treeview(root, columns=("Name", "Match Status"), show="headings")
        self.tree.heading("Name", text="Resort Name")
        self.tree.heading("Match Status", text="Match Status")
        self.tree.column("Name", width=150)
        self.tree.column("Match Status", width=100)
        # First, pack the TreeView
        self.tree.pack(expand=True, fill='both', side='top')
        
        # Then, create and pack the scrollbar
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill='y')
        self.tree.configure(yscrollcommand=scrollbar.set)


        self.match_status_frame = ttk.Frame(root)
        self.match_status_frame.pack(fill='x', side='bottom')
        self.matched_resorts_label = ttk.Label(self.match_status_frame, text="Matched Resorts: 0")
        self.matched_resorts_label.pack(side="left")
        self.no_match_found_label = ttk.Label(self.match_status_frame, text="No Match Found: 0")
        self.no_match_found_label.pack(side="left")
        self.total_processed_label = ttk.Label(self.match_status_frame, text="Total Processed: 0 / 0")
        self.total_processed_label.pack(side="left")

        self.filename = ""
        self.results = None
        self.df_our_resorts = None
        self.df_to_match = None

    def load_excel(self):
        self.filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.filename:
            self.progress['value'] = 0
            self.root.update_idletasks()
            threading.Thread(target=self.process_file, daemon=True).start()

    def process_file(self):
        matched_resorts = 0
        no_match_found = 0
        total_steps = 0
        match_results = []

        try:
            self.df_our_resorts = pd.read_excel(self.filename, sheet_name=0)
            self.df_to_match = pd.read_excel(self.filename, sheet_name=1)
            total_steps = len(self.df_our_resorts.index)

            start_time = time.time()

            self.root.after(0, lambda: self.total_processed_label.config(text=f"Total Processed: 0 / {total_steps}"))
            self.update_progress(10)

            for i, (_, resort_data) in enumerate(self.df_our_resorts.iterrows(), start=1):
                if self.cancel_process:
                    return

                match_result = self.find_best_match(
                    resort_data["Match Status"], resort_data["Match Probability %"], resort_data["Resort Name"], resort_data["Resort Street"], resort_data["Resort City"],
                    resort_data["Resort State"], resort_data["Resort Zip"], resort_data.get("Latitude", "N/A"),
                    resort_data.get("Longitude", "N/A"), self.df_to_match
                )

                elapsed_time = time.time() - start_time
                avg_time_per_item = elapsed_time / i if i > 0 else 0
                estimated_time_remaining = avg_time_per_item * (total_steps - i) if i < total_steps else 0
                
                # Update progress including estimated time remaining
                self.update_progress(20 + (80 * i // total_steps), estimated_time_remaining)  # Adjust progress calculation
                self.processed_rows_label.config(text=f"Rows Processed: {i}/{total_steps}")

                if match_result and match_result[0]['Match Status'] == "Matched Resort":
                    # If a matched resort is found, count it as a match
                    matched_resorts += 1
                    tree_value = f"Matched Resort - {match_result[0]['Matching Resort Name']} (ID: {match_result[0]['Matching Resort ID']})"
                else:
                    # If no suitable match is found, count it once as "No Match Found"
                    no_match_found += 1
                    tree_value = "No Match Found"

                self.tree.insert("", "end", values=(resort_data["Resort Name"], tree_value))


                self.update_progress(20 + (80 * i // total_steps), None)
                self.processed_rows_label.config(text=f"Rows Processed: {i}/{total_steps}")

                self.matched_resorts_label.config(text=f"Matched Resorts: {matched_resorts}")
                self.no_match_found_label.config(text=f"No Match Found: {no_match_found}")
                self.total_processed_label.config(text=f"Total Processed: {i} / {total_steps}")

                # Collect all match results
                match_results.append(match_result)

            # Here, ensure match_results is correctly structured for pd.concat
            self.results = pd.concat([self.df_our_resorts.reset_index(drop=True), pd.DataFrame(match_results)], axis=1)

            flattened_matches = []
            for match_set in match_results:
                # Flatten each set of matches for a single resort entry
                flat_match = {}
                for i, match in enumerate(match_set, start=1):
                    for key, value in match.items():
                        flat_match[f"{key} {i}"] = value
                flattened_matches.append(flat_match)

            # Combine the original DataFrame with the flattened match details
            matches_df = pd.DataFrame(flattened_matches)
            reordered_columns = ['Match Status', 'Match Probability %'] + [col for col in matches_df.columns if col not in ['Match Status', 'Match Probability %']]
            self.final_results = matches_df[reordered_columns]

            if not self.cancel_process:
                self.update_progress(100)
                self.save_results_button.pack(pady=10)
                messagebox.showinfo("Process Complete", "File has been processed. You can now save the results.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file: {e}")

    def cancel_operation(self):
        self.cancel_process = True
        self.reset_gui()

    def reset_gui(self):
        self.progress["value"] = 0
        self.time_remaining_label.config(text="")
        self.processed_rows_label.config(text="Rows Processed: 0 / 0")
        self.matched_resorts_label.config(text="Matched Resorts: 0")
        self.no_match_found_label.config(text="No Match Found: 0")
        self.total_processed_label.config(text="Total Processed: 0 / 0")
        messagebox.showinfo("Operation Cancelled", "The operation was cancelled successfully.")
        self.cancel_process = False

    def update_progress(self, value, estimated_time_remaining=None):
        self.progress["value"] = value
        self.root.update_idletasks()
        if estimated_time_remaining:
            mins, secs = divmod(estimated_time_remaining, 60)
            time_str = f"Estimated Time Remaining: {int(mins)}m {int(secs)}s"
            self.time_remaining_label.config(text=time_str)


    def find_best_match(self, name, street, city, state, zip_code, latitude, longitude, df_to_match):
        name_threshold = 80  # Name similarity threshold for "Highest" probability
        address_threshold = 80  # Address similarity threshold for "Highest" probability
        distance_threshold = 10  # Miles for considering a "Highest" match
        matched_distance_threshold = 50  # Miles for considering a "Matched Resort"
        matched_name_similarity_threshold = 60  # Name similarity for considering a "Matched Resort"

        potential_matches = []

        for _, row in df_to_match.iterrows():
            name_similarity = fuzz.ratio(name.lower(), row['Resort Name'].lower())
            address_similarity = fuzz.ratio(f"{street}, {city}, {state}, {zip_code}".lower(), f"{row['Resort Street']}, {row['Resort City']}, {row['Resort State']}, {row['Resort Zip']}".lower()) if street and city and state and zip_code else 0
            distance = float('inf')

            if pd.notnull(row['Latitude']) and pd.notnull(row['Longitude']) and pd.notnull(latitude) and pd.notnull(longitude):
                try:
                    resort_location = (float(row['Latitude']), float(row['Longitude']))
                    input_location = (float(latitude), float(longitude))
                    distance = geodesic(input_location, resort_location).miles
                except ValueError:
                    pass  # Handle conversion failure gracefully

            # Adjust match probability based on new criteria
            if distance < matched_distance_threshold and name_similarity >= matched_name_similarity_threshold:
                match_status = "Matched Resort"
            else:
                match_status = "No Match Found"

             # Determine match probability
            match_probability = "Low"
            if name_similarity > name_threshold and distance < distance_threshold:
                match_probability = "Medium"
            if address_similarity > address_threshold:
                match_probability = "High"
            if name_similarity > name_threshold and address_similarity > address_threshold and distance < distance_threshold:
                match_probability = "Highest"   

            potential_matches.append({
                "Match Status": match_status,  
                "Match Probability %": match_probability,
                "Matching Resort Name": row['Resort Name'],
                "Matching Resort ID": row['Resort ID'],
                "Name Similarity": name_similarity,
                "Address Similarity": address_similarity,
                "Distance": distance,
                "Matching Address": f"{row['Resort Street']}, {row['Resort City']}, {row['Resort State']}, {row['Resort Zip']}",
                "Latitude": row['Latitude'],
                "Longitude": row['Longitude'] 
            })

        # Sort the potential matches based on status, then distance, name similarity, and address similarity
        potential_matches.sort(key=lambda x: (x['Match Status'] == "No Match Found", x['Distance'], -x['Name Similarity'], -x['Address Similarity']))

        # If any resort is a "Matched Resort", return only that one; otherwise, return top 5 "No Match Found"
        matched_resorts = [match for match in potential_matches if match['Match Status'] == "Matched Resort"]
        if matched_resorts:
            return matched_resorts[:1]  # Return only the best matched resort
        else:
            return potential_matches[:5]  # Return top 5 closest matches for "No Match Found"




    def save_results(self):
        if self.final_results is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
            if save_path:
                with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                    self.final_results.to_excel(writer, sheet_name='Processed Results', index=False)
                    # Additional sheets if needed
                    workbook = writer.book
                    worksheet = writer.sheets['Processed Results']

                    # Apply formatting
                    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                    green_fill = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')
                    white_font = Font(color='FFFFFF')

                    for row in worksheet.iter_rows(min_row=2, max_col=1, max_row=worksheet.max_row):
                        for cell in row:
                            if cell.value == "Matched Resort":
                                cell.fill = green_fill
                            else:
                                cell.fill = red_fill
                            cell.font = white_font

                    # Save the workbook to apply the changes
                    workbook.save(save_path)

                messagebox.showinfo("Save Complete", "Results have been saved.")
        else:
            messagebox.showwarning("No Data", "Please process a file first.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ResortMatcherApp(root)
    root.mainloop()