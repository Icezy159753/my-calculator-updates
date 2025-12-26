def run_this_app(working_dir=None):
    print(f'Hello from RenameSheet! Working dir: {working_dir}')
    import tkinter as tk
    app = tk.Tk()
    app.title('Rename Sheet')
    tk.Label(app, text='This is โปรแกรม RenameSheet').pack()
    app.mainloop()