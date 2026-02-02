import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import pandas as pd
import google.generativeai as genai
import os
import json
import re
from datetime import datetime
from sqlalchemy import create_engine, text

# ==========================================
# 🔑 CONFIGURATION FOR API KEY
# ==========================================
API_KEY = "AIzaSyAja7pgLoX81s55kvyr7iH1nFIpdbKsraQ"
REPORTS_FOLDER = "reports"

# 👇👇👇 MYSQL CONFIGURATION 👇👇👇
DB_HOST = "localhost"
DB_USER = ""
DB_PASSWORD = ""
DB_NAME = "timesheet_reports"

# ==========================================
# ⚙️ CORE LOGIC
# ==========================================

def get_available_model():
    """Autodetects the best available model."""
    try:
        genai.configure(api_key=API_KEY)
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                if 'flash' in m.name or 'pro' in m.name:
                    return m.name
        return 'gemini-1.5-flash'
    except:
        return 'gemini-1.5-flash'

def save_to_mysql(df):
    """Saves to MySQL using SQLAlchemy."""
    try:
        connection_str = f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}/{DB_NAME}"
        engine = create_engine(connection_str)
        with engine.connect() as conn: pass 
        
        df_to_save = df.copy()
        df_to_save['Created_At'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df_to_save.to_sql('work_logs', con=engine, if_exists='append', index=False)
        return True, "Saved to MySQL DB"
    except Exception as e:
        err = str(e)
        if "Access denied" in err: return False, "MySQL Error: Wrong Password or Permission denied."
        elif "Unknown database" in err: return False, "MySQL Error: Create DB in Workbench first."
        return False, f"MySQL Error: {err}"

def auto_save_file(df):
    """Auto-saves to Excel file."""
    try:
        if not os.path.exists(REPORTS_FOLDER): os.makedirs(REPORTS_FOLDER)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"Report_{timestamp}.xlsx"
        full_path = os.path.join(REPORTS_FOLDER, filename)
        df.to_excel(full_path, index=False)
        return True, full_path
    except Exception as e:
        return False, str(e)

def read_messy_excel(file_path):
    print(f"Reading file: {file_path}")
    try:
        with open(file_path, 'rb') as f: header = f.read(8)
        is_binary = header.startswith(b'PK') or header.startswith(b'\xD0\xCF\x11\xE0')
    except: is_binary = False

    try: return pd.read_excel(file_path, dtype=str, engine='calamine'), "Calamine"
    except: pass
    try: return pd.read_excel(file_path, dtype=str, engine='openpyxl'), "OpenPyXL"
    except: pass
    try: dfs = pd.read_html(file_path); return dfs[0], "HTML" if dfs else None
    except: pass
    try: return pd.read_csv(file_path, dtype=str), "CSV"
    except: pass
    
    if not is_binary:
        try:
            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            return pd.DataFrame([content], columns=["Raw_Content"]), "Raw Text"
        except Exception as e: raise ValueError(f"Read Error: {e}")
    else:
        raise ValueError("CRITICAL: Corrupted Binary Excel. Save as CSV.")

def process_and_filter_dates(df):
    """
    Standardizes dates to DD/MM/YYYY, sorts them, and removes weekends.
    """
    try:
        # 1. Convert Date column to datetime objects
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
        
        # 2. Drop rows where Date failed to parse
        df = df.dropna(subset=['Date'])
        
        # 3. Filter: Keep only Monday (0) to Friday (4)
        df = df[df['Date'].dt.dayofweek < 5]
        
        # 4. Sort Chronologically
        df = df.sort_values(by='Date')
        
        # 5. Convert to DD/MM/YYYY format
        df['Date'] = df['Date'].dt.strftime('%d/%m/%Y')
        
        return df
    except Exception as e:
        print(f"Warning: Could not sort/filter dates: {e}")
        return df 

def analyze_sheet_with_ai(file_path):
    try:
        model_name = get_available_model()
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel(model_name)
        
        df, engine_used = read_messy_excel(file_path)
        raw_data_string = df.to_string(index=False)[:30000]

        # --- UPDATED PROMPT: "Automation & Manual" + Detailed Summary ---
        prompt = f"""
        You are a Timesheet Administrator.
        Task: Group by DATE, Sum Hours, Summarize Work, Classify Type.
        RAW DATA: {raw_data_string}
        
        Output Requirement: JSON list of objects.
        Keys: Date, Summary_of_Work, Hours_Worked, Work_Type
        
        IMPORTANT RULES:
        1. Date must be strictly in DD/MM/YYYY format.
        2. 'Summary_of_Work' should be a detailed single-sentence summary (approx 20-25 words) capturing specific tasks.
        3. 'Work_Type' must be strictly one of these values:
           - 'Manual' (if only manual testing)
           - 'Automation' (if only automation)
           - 'Automation & Manual' (if both types are present for that date)
        """

        response = model.generate_content(prompt)
        
        try:
            text_response = response.text
        except ValueError:
            print("DEBUG: AI Blocked the response (Safety Filters).")
            return None, "Error: AI blocked the file content.", "", ""

        clean_json = re.sub(r'```json|```', '', text_response).strip()
        if not clean_json:
             return None, "Error: AI returned empty text.", raw_data_string, ""

        try:
            data = json.loads(clean_json)
        except:
            if "[" not in clean_json: clean_json = "[" + clean_json + "]"
            data = json.loads(clean_json)

        # Create DataFrame
        report_df = pd.DataFrame(data)
        
        # Filter Dates (Mon-Fri) & Format as DD/MM/YYYY
        report_df = process_and_filter_dates(report_df)
        
        return report_df, f"Success ({engine_used})", raw_data_string, clean_json

    except Exception as e:
        return None, f"Error: {str(e)}", "", ""

# ==========================================
# 🖥️ GUI APPLICATION
# ==========================================

class AITimesheetApp:
    def __init__(self, master):
        self.master = master
        master.title("AI Timesheet (Automation & Manual)")
        master.geometry("1100x700")
        
        style = ttk.Style()
        style.theme_use('default') 
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", rowheight=25)
        style.map('Treeview', background=[('selected', '#347083')])
        style.configure("Treeview.Heading", background="#1f1f1f", foreground="white", relief="flat")
        
        self.report_df = None
        self.last_raw_input = ""
        self.last_ai_output = ""

        # Controls
        frame_top = ttk.Frame(master, padding=10)
        frame_top.pack(fill='x')

        self.btn_select = ttk.Button(frame_top, text="Select File & Process", command=self.load_file)
        self.btn_select.pack(side='left', padx=5)

        self.btn_debug = ttk.Button(frame_top, text="Debug Info", command=self.show_debug, state='disabled')
        self.btn_debug.pack(side='left', padx=5)

        self.lbl_status = ttk.Label(frame_top, text="Ready.", foreground="blue")
        self.lbl_status.pack(side='left', padx=10)

        self.btn_manual_save = ttk.Button(frame_top, text="Manual Save (Excel)", command=self.manual_save, state='disabled')
        self.btn_manual_save.pack(side='right', padx=5)

        # Table
        frame_bot = ttk.Frame(master, padding=10)
        frame_bot.pack(fill='both', expand=True)

        self.tree = ttk.Treeview(frame_bot, columns=("Date", "Summary", "Hours", "Type"), show='headings')
        self.tree.heading("Date", text="Date"); self.tree.column("Date", width=100)
        self.tree.heading("Summary", text="Summary"); self.tree.column("Summary", width=600)
        self.tree.heading("Hours", text="Hours"); self.tree.column("Hours", width=80, anchor='center')
        self.tree.heading("Type", text="Type"); self.tree.column("Type", width=120)

        sb = ttk.Scrollbar(frame_bot, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

    def load_file(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return

        self.lbl_status.config(text="Processing...", foreground="orange")
        self.master.update()

        result = analyze_sheet_with_ai(file_path)
        
        if result[0] is not None:
            df, msg, raw_input, ai_output = result
            self.report_df = df
            self.last_raw_input = raw_input
            self.last_ai_output = ai_output
            
            # --- AUTO SAVE LOGIC ---
            save_msg = ""
            file_saved, file_res = auto_save_file(df)
            save_msg += f"File: OK" if file_saved else f"File: {file_res}"

            db_saved, db_res = save_to_mysql(df)
            save_msg += f" | DB: {db_res}"

            self.show_data(df)
            
            color = "green" if db_saved else "red"
            self.lbl_status.config(text=f"{msg} | {save_msg}", foreground=color)
            
            self.btn_manual_save.config(state='normal')
            self.btn_debug.config(state='normal')
            
            if not db_saved: messagebox.showerror("MySQL Error", db_res)
        else:
            self.lbl_status.config(text="Failed.", foreground="red")
            messagebox.showerror("Error", result[1])

    def show_data(self, df):
        for row in self.tree.get_children(): self.tree.delete(row)
        if not df.empty:
            df.columns = [c.lower() for c in df.columns]
            for _, row in df.iterrows():
                self.tree.insert("", "end", values=(
                    row.get('date', ''), row.get('summary_of_work', row.get('summary', '')),
                    row.get('hours_worked', row.get('hours', 0)), row.get('work_type', row.get('type', ''))
                ))

    def show_debug(self):
        win = tk.Toplevel(self.master)
        win.title("Debug"); win.geometry("800x600")
        ttk.Label(win, text="Input:").pack(anchor='w')
        t1 = scrolledtext.ScrolledText(win, height=15); t1.pack(fill='both', expand=True)
        t1.insert('1.0', self.last_raw_input[:5000])
        ttk.Label(win, text="Output:").pack(anchor='w')
        t2 = scrolledtext.ScrolledText(win, height=10); t2.pack(fill='both', expand=True)
        t2.insert('1.0', self.last_ai_output)

    def manual_save(self):
        if self.report_df is None: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.report_df.to_excel(path, index=False)
            messagebox.showinfo("Success", "Saved!")

if __name__ == "__main__":
    root = tk.Tk()
    app = AITimesheetApp(root)
    root.mainloop()
