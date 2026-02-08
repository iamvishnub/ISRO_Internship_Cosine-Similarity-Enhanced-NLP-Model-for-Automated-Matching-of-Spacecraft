import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import threading
import pyodbc
import re
import csv
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
#MDB DETAILS -------------------------------------------------------------------------------------------------------------------
DRV = '{Microsoft Access Driver (*.mdb, *.accdb)}'
PWD = 'A2z1TwO9'
#OBTAINING PREFIX---------------------------------------------------------------------------------------------------------------
def get_pin_prefix(PinNo):
    match = re.match(r'^([A-Z]+)', PinNo.upper())
    return match.group(1) if match else ""
#GUI----------------------------------------------------------------------------------------------------------------------------
class PinMatcherApp(tk.Tk):
    COSINE_THRESHOLD=0.85
    def __init__(self):
        super().__init__()
        self.title("PinDetails Matcher")
        self.geometry("800x600")
        self.configure(bg="#2c3e50")
        messagebox.showinfo("Upload MDB File", "Please upload/select your Microsoft Access MDB (.mdb or .accdb) file.")
        self.mdb_file = filedialog.askopenfilename(
            title="Select Microsoft Access MDB File",
            filetypes=[("Microsoft Access Files", "*.mdb *.accdb")])
        if not self.mdb_file:
            messagebox.showwarning("No file selected", "No MDB file was selected. Exiting.")
            self.destroy()
            return  
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("TLabel", background="#2c3e50", foreground="#ecf0f1", font=("Arial", 12))
        style.configure("TButton", font=("Arial", 11, "bold"))
        style.map("TButton",foreground=[('active', '#ecf0f1')],background=[('active', '#2c3e50')])
        style.map("TRadioButton",foreground=[('active', '#2980b9'),('!active','#ecf0f1')],background=[('active', '#bdc3c7'),('!active','#2c3e50')])
        ttk.Label(self, text="Pin Matching Based on Logic :").pack(pady=(15, 5))
        modes_frame = ttk.Frame(self)
        modes_frame.pack()
        self.mode_var = tk.StringVar(value='1')
        ttk.Radiobutton(modes_frame, text="MATCH ALL AT ONCE", variable=self.mode_var, value='1').pack(side='left', padx=20)
        ttk.Radiobutton(modes_frame, text="MATCH COMPONENT-WISE ", variable=self.mode_var, value='2').pack(side='left', padx=20)
        top_bar_frame = tk.Frame(self, bg="#2c3e50")
        top_bar_frame.pack(side='top', fill='x', padx=5, pady=5)
        center_frame = tk.Frame(top_bar_frame, bg="#2c3e50")
        center_frame.pack(side='top', pady=5)
        top_right_frame = ttk.Frame(center_frame, style="Blue.TFrame")
        top_right_frame.pack(side='left', fill=None, expand=False)
        self.lookup_mdb_files = []
        self.btn_select_lookup = ttk.Button(top_right_frame, text="Select Lookup DB",command=self.select_lookup_db,width=15, style='Compact.TButton')
        self.btn_select_lookup.pack(side='left', padx=2, pady=2)
        self.lookup_status_label = ttk.Label(top_right_frame,text="Lookup DB: No lookup table selected",style="Blue.TLabel",font=("Consolas", 11, "italic"))
        self.lookup_status_label.pack(side='left', padx=(15, 0), pady=2)
        button_frame = ttk.Frame(self)
        button_frame.pack(pady=2)
        self.btn_clear_destination = ttk.Button(button_frame, text="Clear Destination", command=self.clear_destination)
        self.btn_clear_destination.pack(side='left', padx=15)        
        self.btn_run = ttk.Button(button_frame, text="Run Matching", command=self.run_matching_thread)
        self.btn_run.pack(side='left', padx=15)
        self.btn_export_csv = ttk.Button(button_frame, text="Export CSV", command=self.export_csv)
        self.btn_export_csv.pack(side='left', padx=15)
        self.btn_export_mdb = ttk.Button(button_frame, text="Update MDB", command=self.export_mdb)
        self.btn_export_mdb.pack(side='left', padx=15)
        self.btn_close = ttk.Button(button_frame, text="Close", command=self.on_closing)
        self.btn_close.pack(side='left', padx=15)
        text_frame=ttk.Frame(self)
        text_frame.pack(fill='both',padx=15,pady=(0,15),expand=True)
        self.log_text=tk.Text(text_frame,height=20,bg="#2c3e50", fg="#ecf0f1", font=("Consolas", 11),wrap='none')
        self.log_text.pack(side='left',fill='both',expand=True)
        v_scrollbar=ttk.Scrollbar(text_frame,orient='vertical',command=self.log_text.yview)
        v_scrollbar.pack(side='right',fill='y')
        self.log_text.configure(yscrollcommand=v_scrollbar.set)
        h_text_frame=ttk.Frame(self)
        h_text_frame.pack(fill='both',padx=(15,15),pady=1,expand=False)
        h_scrollbar=ttk.Scrollbar(h_text_frame,orient='horizontal',command=self.log_text.xview)
        h_scrollbar.pack(side='bottom',fill='x')
        self.log_text.configure(xscrollcommand=h_scrollbar.set)
        right_side_frame=ttk.Frame(style="Blue.TFrame")
        right_side_frame.pack(side='right',fill=None,padx=10,pady=5)
        self.conn = None
        self.crsr = None
        self.all_data = None
        self.column_names = None
        self.lookup_file = None
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update()
#reverse update query---------------------------------------------------------------------------------------------------------
    def update_reverse_destination(self,from_pin,to_pin):
        """Set reverse Destination: if from_pin's Destination is to_pin, set to_pin's Destination as from_pin unless it's NC or - ."""
        if not to_pin or to_pin.upper() in {"-","NC"}:
            return
        try:
            self.crsr.execute("SELECT Destination FROM PinDetails WHERE PinNo = ? ",(to_pin))
            row=self.crsr.fetchone()
            if row is None :
                return
            current_dest=row[0]
            if current_dest != from_pin:
                self.crsr.execute("UPDATE PinDetails SET Destination = ? WHERE PinNo = ?",(from_pin,to_pin))
        except Exception as e:
            self.log(f"error updating ({to_pin}): {e}")
#select lookup db------------------------------------------------------------------------------------------------
    def select_lookup_db(self):
        files = filedialog.askopenfilenames(
            title="Select Lookup MDB Files",
            filetypes=[("Microsoft Access Files", "*.mdb *.accdb")]
        )
        if files:
            self.lookup_mdb_files = list(files)  # Store list of selected files
            self.lookup_status_label.config(text=f"{len(self.lookup_mdb_files)} lookup files selected")
            self.log(f"Selected lookup MDB files: {self.lookup_mdb_files}")
        else:
            self.lookup_mdb_files = []
            self.lookup_status_label.config(text="Lookup DB: No lookup table selected")
            self.log("No lookup database selected.")

#MATCHING----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    def run_matching_thread(self):
        mode = self.mode_var.get()
        if mode == '2':
            prefix = simpledialog.askstring("Input", "Enter Component For Matching : ", parent=self)
            if prefix is None:
                self.log("Matching cancelled by user.")
                return
            prefix = prefix.strip().upper()
            if not prefix:
                self.log("No prefix entered; aborting matching.")
                messagebox.showwarning("Input Required", "You must enter a component prefix for matching.")
                return
            threading.Thread(target=self.run_matching, args=(prefix,), daemon=True).start()
        else:
            threading.Thread(target=self.run_matching, daemon=True).start()
    def run_matching(self, prefix=None):
        self.log(f"Connecting to database: {self.mdb_file}")
        try:
            self.conn = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, self.mdb_file, PWD))
            self.crsr = self.conn.cursor()
        except Exception as e:
            self.log(f"ERROR: Cannot connect to MDB: {e}")
            messagebox.showerror("DB Connection Error", f"Failed to connect to MDB database:\n{e}")
            return
        mode = self.mode_var.get()
        if mode == '1':
            self.match_all()
            # After matching all pins, count globally empty destinations (all pins)
            try:
                self.crsr.execute("SELECT COUNT(*) FROM PinDetails WHERE Destination = '' ")
                count_empty = self.crsr.fetchone()[0]
            except Exception as e:
                self.log(f"Error checking destination {e}")
                count_empty = 0
        
        
        
            if count_empty > 0:
                self.log(f"\n{count_empty} records have empty destinations fields.")
                answer = messagebox.askyesno(
                    "Empty found",
                    f"There are {count_empty} pins not connected. Do you want to use lookup based approach?")
                if answer:
                    if self.lookup_mdb_files:
                        for lookup_file in self.lookup_mdb_files:
                            self.update_empty_destinations_from_matched_db(lookup_file)
                    else:
                        self.log("No lookup MDB files selected for updating empty destinations.")

        
        
        
        
        
        else:
            if prefix is None:
                self.log("No prefix provided. Aborting.")
                return
            self.match_component(prefix)
            # After matching only the pins starting with prefix, count empty only for those pins
            try:
                self.crsr.execute("SELECT COUNT(*) FROM PinDetails WHERE Destination = '' AND PinNo LIKE ?", (prefix + '%',))
                count_empty = self.crsr.fetchone()[0]
            except Exception as e:
                self.log(f"Error checking destination {e}")
                count_empty = 0
            if count_empty > 0:
                self.log(f"\n{count_empty} records with prefix '{prefix}' have empty destinations fields")
                answer = messagebox.askyesno("Empty found",
                                            f"There are {count_empty} pins starting with '{prefix}' not connected. Do you want to use lookup based approach?")
                
                
                
                if answer:
                    if self.lookup_mdb_files:
                        for lookup_file in self.lookup_mdb_files:
                            self.update_empty_destinations_from_matched_db(lookup_file, prefix=prefix)
                    else:
                        self.log("no path provided")
                        
                        
                        
        self.log("\nMatching process completed.")
#MATCHING ALL AT ONCE ---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
    def match_all(self):
#         self.log("Running : Match all at once.")
#         self.crsr.execute('SELECT * FROM PinDetails')
#         self.all_data = self.crsr.fetchall()
#         self.column_names = [col[0] for col in self.crsr.description]
#         Pinno_index = self.column_names.index("PinNo")
#         desc_index = self.column_names.index("FunlDescription")
#         all_descriptions = [row[desc_index] if row[desc_index] else "" for row in self.all_data]
#         vectorizer = TfidfVectorizer().fit(all_descriptions)
#         desc_vectors = vectorizer.transform(all_descriptions)
#         self.crsr.execute("SELECT Destination FROM PinDetails")
#         used_destinations = set()
#         for row in self.crsr.fetchall():
#             val = row[0]
#             if val and val.upper() not in ["NC", "-"]:
#                 used_destinations.add(val.upper())
#         for idx, row in enumerate(self.all_data):
#             Pin_No = row[Pinno_index]
#             description = row[desc_index].strip() if row[desc_index] else ""
#             description_lower = description.lower()
#             destination_value = None
#             self.log(f"Processing PinNo: {Pin_No} - Description: {description}")
#             if description == "-":
#                 destination_value = "-"
#             elif "nc" in description_lower :
#                 destination_value = "NC"
#             elif description_lower==("%"+"nc"+"%"):
#                 destination_value ="NC"
#             else:
#                 patterns = [r"connect to ([\w\.]+)",r"conn to ([\w\.]+)",r"connecto to ([\w\.]+)",r"connectto to ([\w\.]+)",r"connec to ([\w\.]+)"]
#                 for pattern in patterns:
#                     match = re.search(pattern, description_lower)
#                     if match:
#                         destination_value = match.group(1).upper()
#                         break
#             if destination_value is None:
#                 exact_match_pin=None
#                 for candidate_row in self.all_data:
#                     candidate_Pinno=candidate_row[Pinno_index]
#                     candidate_prefix = get_pin_prefix(candidate_Pinno)
#                     input_pin_prefix = get_pin_prefix(Pin_No)
#                     if(candidate_Pinno.upper()!=Pin_No.upper() and candidate_prefix!=input_pin_prefix and candidate_Pinno.upper() not in used_destinations):
#                         exact_match_pin=candidate_Pinno
#                         break
#             if exact_match_pin:
#                 destination_value=exact_match_pin
#                 self.log(f"Full match found -> Destination{exact_match_pin}")
#             if not destination_value:
#                 desc_vec = vectorizer.transform([description])
#                 similarities = cosine_similarity(desc_vec, desc_vectors)[0]
#                 best_score = 0
#                 best_Pinno = None
#                 input_pin_prefix = get_pin_prefix(Pin_No)
#                 for i, score in enumerate(similarities):
#                     candidate_Pinno = self.all_data[i][Pinno_index]
#                     candidate_prefix = get_pin_prefix(candidate_Pinno)
#                     if(candidate_Pinno.upper()!=Pin_No.upper() and candidate_prefix!=input_pin_prefix and candidate_Pinno.upper() not in used_destinations and score>best_score):
#                         best_score = score
#                         best_Pinno = candidate_Pinno
#                 if best_score >= self.COSINE_THRESHOLD and best_Pinno:
#                     destination_value = best_Pinno
#                     self.log(f"  Cosine Match found (score = {best_score:.2f}) -> Destination : {best_Pinno}")
#                 else:
#                     destination_value = ""
#                     self.log("  No strong match found")
#             if destination_value and destination_value not in ["NC",'-',""]:
#                 used_destinations.add(destination_value.upper())
#             update_query = "UPDATE PinDetails SET Destination = ? WHERE PinNo = ?"
#             try:
#                 self.crsr.execute(update_query, (destination_value, Pin_No))
#                 self.update_reverse_destination(Pin_No,destination_value)
#             except Exception as e:
#                 self.log(f"ERROR updating PinNo {Pin_No}: {e}")
#         self.conn.commit()
#         self.log("Matching and update completed.")
        
        
        
        self.log("Running Mode 1: Match all at once.")
        self.crsr.execute('SELECT * FROM PinDetails')
        self.all_data = self.crsr.fetchall()
        self.column_names = [col[0] for col in self.crsr.description]
        Pinno_index = self.column_names.index("PinNo")
        desc_index = self.column_names.index("FunlDescription")
        all_descriptions = [row[desc_index] if row[desc_index] else "" for row in self.all_data]
        vectorizer = TfidfVectorizer().fit(all_descriptions)
        desc_vectors = vectorizer.transform(all_descriptions)
        self.crsr.execute("SELECT Destination FROM PinDetails")
        used_destinations = set()
        for row in self.crsr.fetchall():
            val = row[0]
            if val and val.upper() not in ["NC", "-"]:
                used_destinations.add(val.upper())
        for idx, row in enumerate(self.all_data):
            Pin_No = row[Pinno_index]
            description = row[desc_index].strip() if row[desc_index] else ""
            description_lower = description.lower()
            destination_value = None
            self.log(f"Processing PinNo: {Pin_No} - Description: {description}")
            if description == ".":
                destination_value = "-"
            elif description_lower.startswith("nc"):
                destination_value = "NC"
            else:
                patterns = [r"connect to ([\w\.]+)",r"conn to ([\w\.]+)",r"connecto to ([\w\.]+)",r"connectto to ([\w\.]+)"]
                for pattern in patterns:
                    match = re.search(pattern, description_lower)
                    if match:
                        destination_value = match.group(1).upper()
                        break
            if not destination_value:
                desc_vec = vectorizer.transform([description])
                similarities = cosine_similarity(desc_vec, desc_vectors)[0]
                best_score = 0
                best_Pinno = None
                input_pin_prefix = get_pin_prefix(Pin_No)
                for i, score in enumerate(similarities):
                    candidate_Pinno = self.all_data[i][Pinno_index]
                    candidate_prefix = get_pin_prefix(candidate_Pinno)
                    if (candidate_Pinno.upper() != Pin_No.upper() and
                        candidate_prefix == input_pin_prefix and
                        candidate_Pinno.upper() not in used_destinations and
                        score > best_score):
                        best_score = score
                        best_Pinno = candidate_Pinno
                if best_score >= 0.60 and best_Pinno:
                    destination_value = best_Pinno
                    self.log(f"  Cosine Match found (score = {best_score:.2f}) -> Destination : {best_Pinno}")
                else:
                    destination_value = ""
                    self.log("  No strong match found")
            update_query = "UPDATE PinDetails SET Destination = ? WHERE PinNo = ?"
            try:
                self.crsr.execute(update_query, (destination_value, Pin_No))
            except Exception as e:
                self.log(f"ERROR updating PinNo {Pin_No}: {e}")
        self.conn.commit()
        self.log("Mode 1 matching and update completed.")        
#MATCHING COMPONENT BY COMPONENT ----------------------------------------------------------------------------------------------------------------------------------------------------
    def match_component(self, Partial_PinNo):
        self.log(f"Running: Match component by component for prefix '{Partial_PinNo}'.")
        self.crsr.execute('SELECT * FROM PinDetails')
        self.all_data = self.crsr.fetchall()
        self.column_names = [col[0] for col in self.crsr.description]
        Pinno_index = self.column_names.index("PinNo")
        desc_index = self.column_names.index("FunlDescription")
        all_descriptions = [row[desc_index] if row[desc_index] else "" for row in self.all_data]
        vectorizer = TfidfVectorizer().fit(all_descriptions)
        desc_vectors = vectorizer.transform(all_descriptions)
        self.crsr.execute("SELECT Destination FROM PinDetails")
        used_destinations = set()
        for row in self.crsr.fetchall():
            val = row[0]
            if val and val.upper() not in ["NC", "-"]:
                used_destinations.add(val.upper())
        self.crsr.execute("SELECT PinNo, FunlDescription FROM PinDetails WHERE PinNo LIKE ?", (Partial_PinNo + '%',))
        input_rows = self.crsr.fetchall()
        if not input_rows:
            self.log("No matching Pin found")
            return
        else:
            self.log(f"PinNo. starting with {Partial_PinNo}")
            prefix_to_remove = Partial_PinNo.lower()
            for Pin_row in input_rows:
                Pin_No = Pin_row[0]
                description = Pin_row[1].strip() if Pin_row[1] else ""
                description_lower = description.lower()
                if description_lower.startswith(prefix_to_remove):
                    description_lower = description_lower[len(prefix_to_remove):]
                destination_value = None
                self.log(f"{Pin_No} - (description)")
                if description == "-":
                    destination_value = "-"
                elif "nc" in description_lower:
                    destination_value = "NC"
                elif description_lower == ("%"+"nc"+"%"):
                    destination_value="NC"
                else:
                    patterns = [r"connect to ([\w\.]+)",r"conn to ([\w\.]+)",r"connecto to ([\w\.]+)",r"connectto to ([\w\.]+)",r"connec to ([\w\.]+)"]
                    for pattern in patterns:
                        match = re.search(pattern, description_lower)
                        if match:
                            destination_value = match.group(1).upper()
                            break
                if not destination_value:
                    desc_vec = vectorizer.transform([description])
                    similarities = cosine_similarity(desc_vec, desc_vectors)[0]
                    best_score = 0
                    best_Pinno = None
                    input_pin_prefix = get_pin_prefix(Pin_No)
                    for i, score in enumerate(similarities):
                        candidate_Pinno = self.all_data[i][Pinno_index]
                        candidate_prefix = get_pin_prefix(candidate_Pinno)
                        if(candidate_Pinno.upper()!=Pin_No.upper() and candidate_prefix!=input_pin_prefix and candidate_Pinno.upper() not in used_destinations and score>best_score):
                            best_score = score
                            best_Pinno = candidate_Pinno
                    if best_score >= self.COSINE_THRESHOLD and best_Pinno:
                        destination_value = best_Pinno
                        self.log(f"\n Cosine Match found (score = {best_score:.2f}) -> Destination : {best_Pinno} \n")
                    else:
                        destination_value = ""
                        self.log("\n No strong match found")
                if destination_value and destination_value not in ["NC",'-',""]:
                    used_destinations.add(destination_value.upper())
                update_query = "UPDATE PinDetails SET Destination = ? WHERE PinNo = ?"
                try:
                    self.crsr.execute(update_query, (destination_value, Pin_No))
                    self.update_reverse_destination(Pin_No, destination_value)
                except Exception as e:
                    self.log(f"ERROR updating PinNo {Pin_No}: {e}")
            self.conn.commit()
            self.log("Matching and update completed.")
#clear destination--------------------------------------------------------------------------------------------------------------
    def clear_destination(self):
        try:
            self.conn = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, self.mdb_file, PWD))
            self.crsr = self.conn.cursor()
        except Exception as e:
            self.log(f"ERROR: Cannot connect to MDB: {e}")
            messagebox.showerror("DB Connection Error", f"Failed to connect to MDB database:\n{e}")
            return
        if not self.conn or not self.crsr:
            messagebox.showwarning("Warning", "mdb file not connected or all values null")
            return
        confirm = messagebox.askyesno("confirm","Are you sure you want to remove all the values in the destination field ?")
        if not confirm:
            self.log("cancelled")
            return
        try:
            self.crsr.execute("UPDATE PinDetails SET Destination = '' ")
            self.conn.commit()
            self.log("All destinations are cleared ")
        except Exception as e:
            self.log(f"Error :{e}") 
    def update_empty_destinations_from_matched_db(self, lookup_file, prefix=None):
        try:
            matched_conn = pyodbc.connect(f'DRIVER={DRV};DBQ={lookup_file};PWD={PWD}')
            matched_crsr = matched_conn.cursor()
        except Exception as e:
            self.log(f"Error opening lookup DB '{lookup_file}': {e}")
            return

        self.log(f"Updating empty Destinations from lookup table: {lookup_file}")
        
        # Query differs based on mode (presence of prefix)
        if prefix:
            # Only pins with given prefix and empty or NULL Destination
            query = "SELECT PinNo, FunlDescription FROM PinDetails WHERE (Destination = '' OR Destination IS NULL) AND PinNo LIKE ?"
            self.crsr.execute(query, (prefix + '%',))
        else:
            # Mode 1: All pins with empty or NULL Destination
            query = "SELECT PinNo, FunlDescription FROM PinDetails WHERE Destination = '' OR Destination IS NULL"
            self.crsr.execute(query)
        
        rows_to_update = self.crsr.fetchall()
        updated_count = 0

        for pin_no, funl_desc in rows_to_update:
            matched_crsr.execute("SELECT Destination FROM PinDetails WHERE PinNo = ? AND FunlDescription = ?", (pin_no, funl_desc))
            matched_row = matched_crsr.fetchone()

            if matched_row:
                matched_destination = matched_row[0]
                if matched_destination and matched_destination.strip():
                    try:
                        self.crsr.execute("UPDATE PinDetails SET Destination = ? WHERE PinNo = ?", (matched_destination, pin_no))
                        updated_count += 1
                        self.log(f"Updated PinNo {pin_no} with Destination '{matched_destination}' from lookup database")
                    except Exception as e:
                        self.log(f"Error updating PinNo {pin_no}: {e}")

        self.conn.commit()
        matched_conn.close()
        self.log(f"Finished updating. Total updated: {updated_count}")
        messagebox.showinfo("Update Complete", f"Updated {updated_count} empty destinations from lookup table")

#EXPORTING AS CSV FILE --------------------------------------------------------------------------------------------------------------------------------------------------------------
    def export_csv(self):
        if not self.conn or not self.crsr:
            messagebox.showwarning("Warning", "Run the matching process before exporting.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return
        try:
            self.crsr.execute('SELECT * FROM PinDetails')
            data = self.crsr.fetchall()
            columns = [col[0] for col in self.crsr.description]
            with open(file_path, 'w', newline='', encoding='utf_8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(columns)
                for row in data:
                    writer.writerow(row)
            self.log(f"CSV exported successfully to: {file_path}")
            messagebox.showinfo("Export Complete", f"CSV exported successfully to:\n{file_path}")
        except Exception as e:
            self.log(f"Error exporting CSV: {e}")
            messagebox.showerror("Export Error", f"Failed to export CSV:\n{e}")
#UPDATING THE MDB FILE --------------------------------------------------------------------------------------------------------------------------------------------------------------
    def export_mdb(self):
        if not self.conn or not self.crsr:
            messagebox.showwarning("Warning", "Run the matching process before exporting.")
            return
        self.log("Starting MDB update...")
        try:
            self.crsr.execute('SELECT * FROM PinDetails')
            updated_data = self.crsr.fetchall()
            column_names = [col[0] for col in self.crsr.description]
            update_query = f'''UPDATE [PinDetails] SET {", ".join([f"[{col}] = ?" for col in column_names if col != "PinNo"])} WHERE [PinNo] = ?'''
            placeholders = ','.join(['?'] * len(column_names))
            insert_query = f"INSERT INTO [PinDetails] ({', '.join([f'[{col}]' for col in column_names])}) VALUES ({placeholders})"
            self.crsr.execute("SELECT [PinNo] FROM [PinDetails]")
            existing_pinnos = set(row[0] for row in self.crsr.fetchall())
            for row in updated_data:
                PinNo_value = row[column_names.index('PinNo')]
                if PinNo_value in existing_pinnos:
                    update_values = [row[column_names.index(col)] for col in column_names if col != "PinNo"] + [PinNo_value]
                    self.crsr.execute(update_query, update_values)
                else:
                    self.crsr.execute(insert_query, row)
            self.conn.commit()
            self.log("MDB file updated successfully.")
            messagebox.showinfo("Export Complete", "MDB file updated successfully.")
        except Exception as e:
            self.log(f"Error updating MDB: {e}")
            messagebox.showerror("Export Error", f"Failed to update MDB:\n{e}")
#CLOSING-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    def on_closing(self):
        try:
            if self.crsr is not None:
                self.crsr.close()
                self.crsr = None
            if self.conn is not None:
                self.conn.close()
                self.conn = None
        except Exception as e:
            self.log(f"Error closing database connection: {e}")
        self.destroy()

if __name__ == "__main__": 
    app = PinMatcherApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()