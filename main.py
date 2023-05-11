import tkinter as tk
import os



# Script for the GUI of the user name input


window = tk.Tk()


label = tk.Label(window, text="Plesa write your name: ")
label.pack()

entry = tk.Entry(window)
entry.pack()



def save_input():
  # Function which saves the input of an user

    user_input = entry.get()
    with open("input.txt", "w") as file:
        file.write(user_input)

def finish():
    save_input()
    window.destroy()
    




button = tk.Button(window, text="Submit", command=finish)
button.pack()
# Run the main event loop


window.mainloop()


# Run program which automates procedure
with open("input.txt","r") as file:
    nazwa = file.readlines()
    if str(nazwa[0]) == "me":
        import rotation_fixer_for_me
    else:
        import rotation_fixer_others


f = open('input.txt', 'r+')
f.truncate(0)
f.close()
os.remove('input.txt')


if __name__ == "__main__":
      
    rotation_fixer()
    


