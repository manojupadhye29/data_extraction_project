# import tkinter
#
# ssw = tkinter.Tk()
#
# def six():
#     toplvl = tkinter.Toplevel() #created Toplevel widger
#     photo = tkinter.PhotoImage(file = 'loading.gif')
#     lbl = tkinter.Label(toplvl ,image = photo)
#     lbl.image = photo #keeping a reference in this line
#     lbl.grid(row=0, column=0)
#
# def base():
#     la = tkinter.Button(ssw,text = 'yes',command=six)
#     la.grid(row=0, column=0) #specifying row and column values is much better
#
# base()

# ssw.mainloop()

# from tkinter import *
# import time
# import os
# root = Tk()
#
# frameCnt = 12
# frames = [PhotoImage(file='loading.gif',format = 'gif -index %i' %(i)) for i in range(frameCnt)]
#
# def update(ind):
#     frame = frames[ind]
#     ind += 1
#     if ind == frameCnt:
#         ind = 0
#     label.configure(image=frame)
#     root.after(100, update, ind)
# label = Label(root)
# label.pack()
# root.after(0, update, 0)
# root.mainloop()


# import tkinter as tk
# from time import sleep
#
# def task():
#     # The window will stay open until this function call ends.
#     sleep(2) # Replace this with the code you want to run
#     root.destroy()
#
# root = tk.Tk()
# root.title("Example")
#
# label = tk.Label(root, text="Waiting for task to finish.")
# label.pack()
#
# root.after(200, task)
# root.mainloop()
#
# print("Main loop is now over and we can do other stuff.")
from time import sleep

# import customtkinter
#
# class ToplevelWindow(customtkinter.CTkToplevel):
#     def __init__(self, *args, **kwargs):
#         super().__init__(*args, **kwargs)
#         self.geometry("400x300")
#
#         self.label = customtkinter.CTkLabel(self, text="ToplevelWindow")
#         self.label.pack(padx=20, pady=20)
#
# class App(customtkinter.CTk):
#     def __init__(self, *args, **kwargs):
#         super().__init__(*args, **kwargs)
#         self.geometry("500x400")
#
#         self.button_1 = customtkinter.CTkButton(self, text="open toplevel", command=self.open_toplevel)
#         self.button_1.pack(side="center", padx=20, pady=20)
#
#         self.toplevel_window = None
#
#     def open_toplevel(self):
#         if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
#             self.toplevel_window = ToplevelWindow(self)  # create window if its None or destroyed
#         else:
#             self.toplevel_window.destroy()
#
#
# if __name__ == "__main__":
#     app = App()
#     app.mainloop()

import tkinter as tk
from PIL import Image, ImageTk

class GifImage(tk.Canvas):
    def __init__(self, parent, gif_file, *args, **kwargs):
        tk.Canvas.__init__(self, parent, *args, **kwargs)

        # load the animated GIF using the PIL module
        self.gif = Image.open(gif_file)
        self.frames = []
        try:
            while True:
                self.frames.append(ImageTk.PhotoImage(self.gif.copy()))
                self.gif.seek(len(self.frames)) # move to next frame
        except EOFError:
            pass

        # create a Label widget to display the animated GIF
        self.current_frame = 0
        self.label = tk.Label(self, image=self.frames[0])
        self.label.pack()

        # schedule the animation
        self.after(0, self.animate)

    def animate(self):
        self.label.configure(image=self.frames[self.current_frame])
        self.current_frame = (self.current_frame + 1) % len(self.frames)
        self.after(50, self.animate)

# create the main window and the custom widget
root = tk.Tk()
gif_widget = GifImage(root, "loading.gif")
gif_widget.pack()

# run the main window
root.mainloop()




# import tkinter as tk
#
# class App:
#     def __init__(self, master):
#         self.master = master
#         master.title("Loading Screen Example")
#
#         # create a button to trigger the loading screen
#         self.button = tk.Button(master, text="Load Data", command=self.load_data)
#         self.button.pack()
#
#     def load_data(self):
#         # disable the button to prevent multiple clicks
#         self.button.config(state="disabled")
#
#         # create a new Toplevel window to show the loading message
#         self.loading_window = tk.Toplevel(self.master)
#         self.loading_window.title("Loading...")
#         self.loading_window.geometry("200x100")
#
#         # create a Label widget to show the loading message
#         self.loading_label = tk.Label(self.loading_window, text="Please wait, loading data...")
#         self.loading_label.pack(pady=20)
#
#         # simulate the loading process by using the after() method to schedule the
#         # update() function to be called after a delay of 3 seconds
#         self.master.after(3000, self.update)
#
#     def update(self):
#         # hide the loading window and enable the button again
#         self.loading_window.destroy()
#         self.button.config(state="normal")
#
#         # do some work here to load the data
#         # ...
#
#         # show a message box to indicate that the data has been loaded
#         tk.messagebox.showinfo("Loading Screen Example", "Data loaded successfully!")
#
# # create the main window and the application instance
# root = tk.Tk()
# app = App(root)
#
# # run the main window
# root.mainloop()

