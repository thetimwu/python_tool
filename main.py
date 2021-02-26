import pyautogui
import tkinter as tk
import requests
import bs4
import re
import time

root=tk.Tk()

root.geometry("460x300")
root.title("Python-screenShot")

le= tk.StringVar()
top= tk.StringVar()
width= tk.StringVar()
height= tk.StringVar()
p = tk.StringVar()
url = tk.StringVar()

tk.Label(root, text = "left").grid(row = 0, column=0, sticky='E')
tk.Entry(root, textvariable=le).grid(row = 0, column = 1, sticky='W')


tk.Label(root, text = "top").grid(row = 1, column=0, sticky='E')
tk.Entry(root, textvariable=top).grid(row = 1, column = 1, sticky='W')


tk.Label(root, text = "width").grid(row = 2, column=0, sticky='E')
tk.Entry(root,textvariable=width).grid(row = 2, column = 1, sticky='W')


tk.Label(root, text = "height").grid(row = 3, column=0, sticky='E')
tk.Entry(root,textvariable=height).grid(row = 3, column = 1, sticky='W')

tk.Label(root, text = "File Path").grid(row = 5, ipady=23)
tk.Entry(root,textvariable=p, width="65").grid(row = 5, column = 1)


tk.Label(root, text = "MP3 URL").grid(row = 9, ipady=40)
tk.Entry(root,textvariable=url, width="65").grid(row = 9, column = 1)

def takes():
    lv = le.get()
    tv = top.get()
    wv = width.get()
    hv = height.get()
    pv = p.get()
    img=pyautogui.screenshot(imageFilename=r""+pv+".png", region=(float(lv),float(tv), float(wv), float(hv)))


def quitApp():
    root.destroy()


def download_mp3():
    pv = p.get()
    pv = pv[:-2]
    uv = url.get()

    for x in range(1, 12):
        new_url = uv.replace('2', str(x))
        print(new_url)
        doc = requests.get(new_url)
        with open('' + pv + str(x) + '.mp3', 'wb') as f:
            f.write(doc.content)


button1=tk.Button(root, text="Take ScreenShot", command=takes)
button1.place(x=60,y=130)

button3=tk.Button(root, text="Download mp3", command=download_mp3)
button3.place(x=60,y=215)

button2=tk.Button(root, text="Quit", command=quitApp)
button2.place(x=410,y=270)

root.mainloop()