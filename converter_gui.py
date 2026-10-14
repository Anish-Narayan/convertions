"""Tkinter front end for the conversion engine (entry point for the .exe)."""
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from converter import convert_sql
from mapping_book import MappingBook


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SQL Pod Converter")
        self.geometry("1200x780")
        self.book, self.last_output = None, None
        self.columnconfigure(0, weight=1); self.columnconfigure(1, weight=1)
        self.rowconfigure(2, weight=3); self.rowconfigure(4, weight=1)

        top = ttk.Frame(self); top.grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=6)
        ttk.Button(top, text="Upload mapping Excel…", command=self.load_book).pack(side="left")
        self.file_lbl = ttk.Label(top, text="No workbook loaded"); self.file_lbl.pack(side="left", padx=10)

        self.src_var, self.tgt_var = tk.StringVar(), tk.StringVar()
        for col, (txt, var) in enumerate((("Source pod", self.src_var), ("Target pod", self.tgt_var))):
            f = ttk.Frame(self); f.grid(row=1, column=col, sticky="ew", padx=8)
            ttk.Label(f, text=txt).pack(side="left")
            cb = ttk.Combobox(f, textvariable=var, state="readonly", width=30); cb.pack(side="left", padx=6)
            setattr(self, "cb_src" if col == 0 else "cb_tgt", cb)
        self.src_txt = scrolledtext.ScrolledText(self, wrap="none", undo=True, font=("Consolas", 10))
        self.out_txt = scrolledtext.ScrolledText(self, wrap="none", undo=True, font=("Consolas", 10))
        self.src_txt.grid(row=2, column=0, sticky="nsew", padx=8, pady=4)
        self.out_txt.grid(row=2, column=1, sticky="nsew", padx=8, pady=4)

        bar = ttk.Frame(self); bar.grid(row=3, column=0, columnspan=2, pady=4)
        for text, cmd in (("Convert", self.convert), ("Clear", self.clear),
                          ("Copy Output", self.copy_out), ("Save SQL", self.save)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=5)
        self.status = scrolledtext.ScrolledText(self, height=10, wrap="word", state="disabled", font=("Consolas", 9))
        self.status.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=8, pady=6)
        self.status.tag_config("err", foreground="#b00020"); self.status.tag_config("ok", foreground="#1b6e1b")

    def log(self, text, tag=None):
        self.status.config(state="normal"); self.status.delete("1.0", "end")
        self.status.insert("end", text, tag); self.status.config(state="disabled")

    def load_book(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            self.book = MappingBook.load(path)
        except Exception as e:  # noqa
            self.book = None
            return self.log(f"Could not read workbook:\n{e}", "err")
        labels = list(self.book.labels)          # populated from the workbook, never hard-coded
        self.cb_src["values"] = self.cb_tgt["values"] = labels
        self.src_var.set(labels[0]); self.tgt_var.set(labels[1])
        self.file_lbl.config(text=path)
        self.log(f"Loaded {len(self.book.table_rows)} table mappings, {len(self.book.column_rows)} column rows.", "ok")

    def convert(self):
        if self.book is None:
            return self.log("Upload a mapping workbook first.", "err")
        if not self.src_var.get() or not self.tgt_var.get() or self.src_var.get() == self.tgt_var.get():
            return self.log("Choose two different pods.", "err")
        sql = self.src_txt.get("1.0", "end").strip()
        force = False
        if self.last_output and sql == self.last_output.strip():
            if not messagebox.askyesno("Already converted", "This text is the previous converted output. Convert again?"):
                return
            force = True
        res = convert_sql(sql, self.book, self.src_var.get(), self.tgt_var.get(), force=force)
        self.out_txt.delete("1.0", "end")
        if res.ok:
            self.out_txt.insert("1.0", res.sql)
            self.last_output = res.sql
        self.log(res.report(), "ok" if res.ok and not res.errors else "err")

    def clear(self):
        self.src_txt.delete("1.0", "end"); self.out_txt.delete("1.0", "end")
        self.last_output = None; self.log("Cleared.")

    def copy_out(self):
        self.clipboard_clear(); self.clipboard_append(self.out_txt.get("1.0", "end").strip())
        self.log("Output copied to clipboard.", "ok")

    def save(self):
        text = self.out_txt.get("1.0", "end").strip()
        if not text:
            return self.log("Nothing to save.", "err")
        path = filedialog.asksaveasfilename(defaultextension=".sql", filetypes=[("SQL", "*.sql")])
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(text + "\n")
            self.log(f"Saved {path}", "ok")


if __name__ == "__main__":
    App().mainloop()
