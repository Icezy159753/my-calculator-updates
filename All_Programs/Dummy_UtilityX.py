def run_this_app(working_dir=None):
    print(f'Hello from Dummy_UtilityX! Working dir: {working_dir}')
    import tkinter as tk
    app = tk.Tk()
    app.title('Utility X')
    tk.Label(app, text='This is Utility X (Dummy)').pack()
    app.mainloop()