import tkinter as tk
from tkinter import filedialog, messagebox
import os
from rule_engine import main as run_engine

class RuleEngineApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Rule Validator")

        self.rule_path = None
        self.doc_path = None
        self.json_path = None

        tk.Button(root, text="Upload Rulebook (Excel)", command=self.upload_rule).pack(pady=5)
        tk.Button(root, text="Upload Document (PDF/Word)", command=self.upload_doc).pack(pady=5)
        tk.Button(root, text="Upload JSON Input", command=self.upload_json).pack(pady=5)
        tk.Button(root, text="Run Validation", command=self.run_validation).pack(pady=20)

    def upload_rule(self):
        self.rule_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.rule_path:
            messagebox.showinfo("Uploaded", f"Rulebook: {os.path.basename(self.rule_path)}")

    def upload_doc(self):
        self.doc_path = filedialog.askopenfilename(filetypes=[("Documents", "*.docx *.pdf")])
        if self.doc_path:
            messagebox.showinfo("Uploaded", f"Document: {os.path.basename(self.doc_path)}")

    def upload_json(self):
        self.json_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if self.json_path:
            messagebox.showinfo("Uploaded", f"JSON: {os.path.basename(self.json_path)}")

    def run_validation(self):
        if not all([self.rule_path, self.doc_path, self.json_path]):
            messagebox.showerror("Error", "Please upload all required files.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not output_path:
            return

        try:
            run_engine(self.rule_path, self.doc_path, self.json_path, output_path)
            messagebox.showinfo("Success", f"Validation complete. Results saved to:{output_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = RuleEngineApp(root)
    root.mainloop()
