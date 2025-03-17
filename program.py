import tkinter as tk

window = tk.Tk()  

window.title("Report Payment Tool")

window_width = 800  
window_height = 600

screen_width = window.winfo_screenwidth() 
screen_height = window.winfo_screenheight() 
x_coordinate = int((screen_width / 2) - (window_width / 2))
y_coordinate = int((screen_height / 2) - (window_height / 2))

window.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
window.minsize(600,400)

window.mainloop()