def run_this_app(working_dir=None):
    print(f'Hello from Var_Value_SPSS_V6! Working dir: {working_dir}')
    import tkinter as tk
    app = tk.Tk()
    app.title('Get Value SPSS')
    tk.Label(app, text='This is Program_Get Value V6').pack()
    app.mainloop()