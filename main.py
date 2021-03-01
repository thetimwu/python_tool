from tkinter import *

window = Tk()
window.minsize(width=500, height=300)

my_label = Label(text="my title", font=("Arial", 24, "bold"))
my_label.pack(side="left")



def handler():
    my_label.config(text=input.get())



button = Button(text='click me', command=handler)
button.pack()


input = Entry(width=10)
input.pack()


window.mainloop()