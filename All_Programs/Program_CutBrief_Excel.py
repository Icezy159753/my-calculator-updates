def run_this_app(working_dir=None):
    print(f'Hello from CutBrief_Excel! Working dir: {working_dir}')
    import tkinter as tk
    app = tk.Tk()
    app.title('Cut Brief')
    tk.Label(app, text='This is โปรแกรมตัดชุด').pack()
    app.mainloop()