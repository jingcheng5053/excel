# import ttk
from tkinter import *
from tkinter import ttk

root = Tk()
tree = ttk.Treeview(root, columns=('col1', 'col2', 'col3'))
tree.column('col1', width=100, anchor='center')
tree.column('col2', width=100, anchor='center')
tree.column('col3', width=100, anchor='center')
tree.heading('col1', text='col1')
tree.heading('col2', text='col2')
tree.heading('col3', text='col3')
tree.pack()

tv = ttk.Treeview(root, height=10, columns=('c1', 'c2', 'c3'))
for i in range(1000):
    tv.insert('', i, values=('a' + str(i), 'b' + str(i), 'c' + str(i)))
tv.pack()
vbar = ttk.Scrollbar(root, orient=VERTICAL, command=tv.yview)
tv.configure(yscrollcommand=vbar.set)
tv.grid(row=0, column=0, sticky=NSEW)
vbar.grid(row=0, column=1, sticky=NS)


def onDBClick(event):
    item = tree.selection()[0]
    print(tree.selection())
    print("you clicked on ", tree.item(item, "values"))


for i in range(10):
    tree.insert('', i, values=('a' + str(i), 'b' + str(i), 'c' + str(i)))
    tree.bind("<Double-1>", onDBClick)

root.mainloop()
