def Center(root):
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth()//2) - (width//2)
    y = (root.winfo_screenheight()//2) - (height//2)
    root.geometry('{}x{}+{}+{}'.format(width,height,x,y))