import customtkinter

# class MyTabView(customtkinter.CTkTabview):
#     def __init__(self, master, **kwargs):
#         super().__init__(master, **kwargs)

#         # create tabs
#         self.add("tab 1")
#         self.add("tab 2")

#         # add widgets on tabs
#         self.label = customtkinter.CTkLabel(master=self.tab("tab 1"))
#         self.label.grid(row=0, column=0, padx=20, pady=10)


# class App(customtkinter.CTk):
#     def __init__(self):
#         super().__init__()

#         self.tab_view = MyTabView(master=self)
#         self.tab_view.grid(row=0, column=0, padx=20, pady=20)


app = customtkinter.CTk()


tabview = customtkinter.CTkTabview(master=app)
tabview.pack(padx=20, pady=20)

tabview.add("tab 1")  # add tab at the end
tabview.add("tab 2")  # add tab at the end
tabview.set("tab 2")  # set currently visible tab

button = customtkinter.CTkButton(master=tabview.tab("tab 1"))
button.pack(padx=20, pady=20)


app.mainloop()