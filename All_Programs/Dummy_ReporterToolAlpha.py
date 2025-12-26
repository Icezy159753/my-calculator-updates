def run_this_app(working_dir=None):
    print(f'Hello from Dummy_ReporterToolAlpha! Working dir: {working_dir}')
    import tkinter as tk
    app = tk.Tk()
    app.title('Reporter Tool Alpha')
    tk.Label(app, text='This is Reporter Tool Alpha (Dummy)').pack()
    app.mainloop()