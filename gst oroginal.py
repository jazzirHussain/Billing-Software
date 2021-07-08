from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from openpyxl import *
from pickle import  *
w = Tk()
w.title('BILL')
w.geometry('800x600')
pickled_in = load_workbook('E:\pickle_n.xlsx')
sheet1 = pickled_in.active
n1 = int(sheet1['A1'].value) + 1
d = 0
flag = 0
flag2 = 0
f = 0
items_var = StringVar()
items_var.set("--SELECT--")
# units
units = ['Sqmtr', 'Kgs', 'Nos', 'Units', 'Cms', 'Mtrs']

# frame
main_frame = Frame(w)
main2_frame = Frame(w)
main3_frame = Frame(w)
main4_frame = Frame(w)
for frame in (main_frame,main2_frame,main3_frame,main4_frame):
    frame.grid (row=0, column=0, sticky=N+E+W+S)
frame_buyer = LabelFrame(main_frame, text="Buyer's Details", padx=10, pady=20, bd=2, relief=GROOVE)
frame_seller = LabelFrame(main_frame, text="Seller's Details", padx=10, pady=20, bd=2, relief=GROOVE)
frame_goods = LabelFrame(main2_frame, text="Good's Details", padx=13, pady=20, bd=2, relief=GROOVE)

btn_goodsframe = Button(main_frame, text='Next', command=lambda: raise_frame(main2_frame),width=10,relief=GROOVE)
btn_goodsframe.grid(row=4,column=3,pady=20)


def raise_frame(frame):
    frame.tkraise()


# labels
lab_intro = Label(main3_frame,text='JNJ Billing Software',font=('GungsuhChe',40),anchor=W)
lab_contact = Label(main3_frame,text='Contact:+91 9496526595',font=('GungsuhChe'),anchor=W)
lab_billno = Label(main_frame, text="Invoice Number: ", anchor=W, pady=10, padx=10)
lab_vhclno = Label(main_frame, text="Vehicle Number: ", anchor=W, pady=10, padx=10)
lab_date = Label(main_frame, text="Invoice Date: ", anchor=W, pady=10, padx=10)
lab_place = Label(frame_buyer, text="Place of supply: ", anchor=E, pady=2, padx=5)
lab_buyername = Label(frame_buyer, text="Buyer's Name: ", anchor=E, padx=5,  pady=5)
lab_buyeradd = Label(frame_buyer, text="Buyer's Address: ", anchor=NE, padx=5, pady=5)
lab_buyertin = Label(frame_buyer, text="Buyer's TIN: ", anchor=E, padx=5, pady=5)
lab_sellername = Label(frame_seller, text="Seller's Name: ", anchor=E, padx=5, pady=5)
lab_selleradd = Label(frame_seller, text="Seller's Address: ", anchor=NE, padx=5, pady=5)
lab_sellertin = Label(frame_seller, text="Seller's TIN: ", anchor=E, padx=5, pady=5)
lab_item = Label(frame_goods, text='Name of Good: ', anchor=E,padx=5, pady=3)
lab_amntitem = Label(frame_goods, text='Amount of Item:', anchor=E,padx=5, pady=3)
lab_priceitem = Label(frame_goods, text='Price of Item per Unit:', anchor=E,padx=5, pady=3)
lab_unit = Label(frame_goods, text='Select Unit', anchor=E,padx=5, pady=3)
lab_tax1 = Label(frame_goods, text='CGST: ', anchor=E,padx=0,  pady=3)
lab_tax2 = Label(frame_goods, text='SGST: ', anchor=E,padx=0,  pady=3)
lab_totalamnt = Label(frame_goods, text='Total Amount: ', anchor=E,padx=0, pady=3)
lab_sellername2= Label(main4_frame, text="Seller's Name: ", anchor=E, padx=5, pady=5)
lab_selleradd2 = Label(main4_frame, text="Seller's Address: ", anchor=NE, padx=5, pady=5)
lab_sellertin2 = Label(main4_frame, text="Seller's TIN: ", anchor=E, padx=5, pady=5)
lab_none = Label(frame_goods, text='        ')
# entries
e_billno = Entry(main_frame)
e_vhclno = Entry(main_frame)
e_date = Entry(main_frame)
e_place = Entry(frame_buyer, width=27)
e_buyername = Entry(frame_buyer, width=27)
e_buyeradd = Text(frame_buyer, width=20, height=5,wrap=WORD)
e_buyertin = Entry(frame_buyer, width=27)
e_sellername = Entry(frame_seller, width=27)
e_sellername.insert(0, 'ZIYAN POLYMERS')
e_selleradd = Text(frame_seller, width=20, height=5,wrap=WORD)
e_selleradd.insert('1.0', 'INDUSTRIAL DEVELOPMENT AREA MUPPATHADOM P O EDAYAR ALUVA')
e_sellertin = Entry(frame_seller, width=27)
e_sellertin.insert(0, '32AVVPS3598R2Z5')
e_item = ttk.Combobox(frame_goods,textvariable=items_var, width=17)
e_amntitem = Entry(frame_goods, width=20)
e_priceitem = Entry(frame_goods, width=20)
e_tax1 = Entry(frame_goods, width=5)
e_tax2 = Entry(frame_goods, width=5)
e_tax1.insert(0, 9)
e_tax2.insert(0, 9)
e_tax1amnt = Entry(frame_goods, width=17)
e_tax2amnt = Entry(frame_goods, width=17)
e_totalamnt = Entry(frame_goods, width=23)
e_sellername2 = Entry(main4_frame, width=27)
e_selleradd2 = Text(main4_frame, width=20, height=5)
e_sellertin2 = Entry(main4_frame, width=27)


# unit dropdown
var = StringVar()
var.set("--SELECT--")
unit_list = ttk.Combobox(frame_goods,textvariable = var, width=17)
# unit list position
unit_list.grid(row=3, column=1)

#combobox

e_item['values'] = ('Plywood','Resin','Veneer')
unit_list['values'] = units
# functions

btns = [e_tax1, e_tax2, e_tax1amnt,e_tax2amnt, e_priceitem,e_amntitem,e_billno,e_vhclno,e_buyertin,
            e_buyername,e_totalamnt]

btn_bill = Button(main3_frame,text='Type Bill',command=lambda:raise_frame(main_frame),width=15,height=2,relief=GROOVE,font=('Narkisim'))
btn_bill.place(relx=.48,rely=.45)




# tree view

tree = ttk.Treeview(main2_frame,columns=('Item','Quanity','Unit','amount'))
tree.grid(row=1,column=0,padx=5,sticky=W,columnspan=8)
tree.column('#1',width=130)
tree.column('#2',width=100)
tree.column('#3',width=150)
tree.column('#4',width=126)
tree.heading('#0',text='Item', anchor=W)
tree.heading('#1',text='Quanity', anchor=W)
tree.heading('#2',text='Unit', anchor=W)
tree.heading('#3',text='Price per unit', anchor=W)
tree.heading('#4',text='Amount', anchor=W)

def btn_add(flag):
    global y
    global d
    global item_quant
    item_price = float (e_priceitem.get ())
    item_quant = float (e_amntitem.get ())
    item_name = e_item.get ()
    item_unit = var.get ()
    item_amnt = item_quant * item_price
    if not e_priceitem.get() or not e_amntitem.get():
        if not e_priceitem.get():
           e_priceitem.insert(0, '0')
        if not e_amntitem.get():
           e_amntitem.insert (0, '0')
        messagebox.showerror ('Error', 'Some fields are empty')
    if flag > d:
        z = tree.index(item_id)
        k = item_id[0]
        tree.delete (item_id)
        tree.insert("", z,k, text=item_name, values=(item_quant, item_unit, item_price, item_amnt))
        for items in y:
            items.delete (0, END)
        var.set('-SELECT-')
        d += 1
    else:
      tree.insert("",END,text=item_name,values=(item_quant,item_unit,item_price,item_amnt))
      for items in y:
          items.delete (0, END)
      var.set ('-SELECT-')

def btn_Reset():
    e_buyeradd.delete('1.0', END)
    var.set ("-SELECT-'")
    items_var.set ("--SELECT--")
    for item in btns:
        item.delete (0, END)
    e_tax1.insert(0, '9')
    e_tax2.insert (0, '9')

def total_amnt():
    global int_amntitem
    global int_priceitem
    global int_tax1
    global int_tax2
    e_totalamnt.delete(0, END)
    e_tax2amnt.delete (0, END)
    e_tax1amnt.delete (0, END)
    if not e_amntitem.get() or not e_priceitem.get():
        e_totalamnt.insert(0, '0')
        e_tax2amnt.insert(0, '0')
        e_tax1amnt.insert (0, '0')
    else:
        int_tax1 = float(e_tax1.get ())
        int_tax2 = float(e_tax2.get ())
        int_amntitem = float(e_amntitem.get ())
        int_priceitem = float(e_priceitem.get ())
        e_totalamnt.delete(0, END)
        e_tax1amnt.delete (0, END)
        e_tax2amnt.delete (0, END)
        e_totalamnt.insert(0, (int_amntitem * int_priceitem) + (
                (int_amntitem * int_priceitem) * (int_tax1 / 100 + int_tax2 / 100)))
        e_tax1amnt.insert(0, (int_amntitem * int_priceitem)* (int_tax1/100))
        e_tax2amnt.insert (0, (int_amntitem * int_priceitem) * (int_tax2/100))


def edit():
    global item_id
    item_id=tree.selection()
    item_name = tree.item(item_id,'text')
    x = tree.item (item_id, 'values')
    item_quant = x[0]
    item_price = x[2]
    item_unit =  x[1]
    global y
    for items in y:
        items.delete(0,END)
    e_item.insert(0,item_name)
    e_amntitem.insert (0, item_quant)
    e_priceitem.insert (0, item_price)
    var.set(item_unit)
    global  flag
    flag += 1



def update():
    e_sellername.delete(0,END)
    e_selleradd.delete('1.0',END)
    e_sellertin.delete(0,END)
    e_sellertin.insert(0, e_sellertin2.get())
    e_sellername.insert(0, e_sellername2.get())
    e_selleradd.insert('1.0',e_selleradd2.get('1.0',END))
    global flag2
    global f
    flag2 += 1
    f += 1


btn_back = Button(main4_frame,text='Back',command=lambda: raise_frame(main3_frame),width=10,relief=GROOVE)
btn_back.place(relx=.51,rely=7)
# buttons
btn_reset = Button(main2_frame, text='Reset', command=btn_Reset, width=10,relief=GROOVE)
btn_get = Button(frame_goods, text='Amnt', width=5, height=1, command=total_amnt,relief=GROOVE)
btn_prnt = Button(main2_frame, text='Print', width=10,relief=GROOVE)
btn_edit = Button(frame_goods, text='Edit',command=edit,width=10,relief=GROOVE)
btn_edit.grid(row=5,column=5,padx=5)
btn_updtadrs = Button(main3_frame,text="Update Adress",command=lambda:raise_frame(main4_frame),width=15,height=2,relief=GROOVE,font=('Narkisim'))
btn_updtadrs.place(relx=.48,rely=.55)
btn_update = Button(main4_frame,text='Update',command=update,width=10,relief=GROOVE)
btn_update.place(relx=.3,rely=.7)
btn_back2 = Button(main4_frame,text='Back',command=lambda: raise_frame(main3_frame),width=10,relief=GROOVE)
btn_back2.place(relx=.4,rely=.7)
# frame position
frame_buyer.grid(row=0, column=0, ipadx=10, ipady=10,padx=10, pady=10, columnspan=2)
frame_seller.grid(row=0, column=2, ipadx=10, ipady=22, columnspan=2)
frame_goods.grid(row=0, column=0,padx=5, pady=10, columnspan=8,sticky=W)



# label position
lab_intro.place(relx=.15,rely=.05)
lab_contact.place(relx=.7,rely=.96)
lab_billno.grid(row=2, column=0, sticky=W+E)
lab_vhclno.grid(row=2, column=2, sticky=W+E)
lab_date.grid(row=3, column=0, sticky=W+E)
lab_place.grid(row=6, column=0, sticky=W+E)
lab_buyername.grid(row=0, column=0, sticky=W+E)
lab_buyeradd.grid(row=1, column=0, sticky=W+E+S+N)
lab_buyertin.grid(row=5, column=0, sticky=W+E)
lab_sellername.grid(row=0, column=0, sticky=W+E)
lab_selleradd.grid(row=1, column=0, sticky=W+E+S+N)
lab_sellertin.grid(row=5, column=0, sticky=W+E)
lab_item.grid(row=0, column=0, sticky=W+E)
lab_amntitem.grid(row=1, column=0, sticky=W+E)
lab_priceitem.grid(row=2, column=0, sticky=W+E)
lab_unit.grid(row=3, column=0, sticky=W+E)
lab_tax1.grid(row=0, column=3, sticky=W+E)
lab_tax2.grid(row=1, column=3, sticky=W+E)
lab_totalamnt.grid(row=2, column=3, sticky=W+E)
lab_none.grid(row=1, column=2)
lab_sellername2.place(relx=.113,rely=.11)
lab_selleradd2.place(relx=.1,rely=.18)
lab_sellertin2.place(relx=.133,rely=.49)
# entry position
e_billno.grid(row=2, column=1)
e_vhclno.grid(row=2, column=3)
e_date.grid(row=3, column=1)
e_buyername.grid(row=0, column=1)
e_buyeradd.grid(row=1, column=1)
e_buyertin.grid(row=5, column=1)
e_sellername.grid(row=0, column=1)
e_selleradd.grid(row=1, column=1)
e_sellertin.grid(row=5, column=1)
e_item.grid(row=0, column=1)
e_amntitem.grid(row=1, column=1)
e_priceitem.grid(row=2, column=1)
e_tax1.grid(row=0, column=3, columnspan=2)
e_tax2.grid(row=1, column=3, columnspan=2)
e_tax1amnt.grid(row=0, column=4, columnspan=3)
e_tax2amnt.grid(row=1, column=4, columnspan=3)
e_totalamnt.grid(row=2, column=4)
e_place.grid(row=6, column=1)
e_sellername2.place(relx=.25,rely=.12)
e_selleradd2.place(relx=.25,rely=.2)
e_sellertin2.place(relx=.25,rely=.5)

# buttons position
btn_reset.grid(row=2, column=6,sticky=W,padx=5)
btn_get.grid(row=2, column=6)
btn_prnt.grid(row=2, column=5,sticky=E)


# excel linking
excel = load_workbook('E:\excel1.xlsx')
sheet = excel.active





btn_back = Button(main2_frame,text='Back',command=lambda: raise_frame(main_frame),width=10,relief=GROOVE)
btn_back.grid(row=2,column=0)


btn_save = Button(main2_frame, text='Save', width=10, command=lambda :save(n1),relief=GROOVE)
btn_save.grid(row=2, column=6,sticky=E,padx=33,pady=10)

def save(n):
    btn_save = Button (main2_frame, text='Save', width=10, command=lambda :save(n+1),relief=GROOVE)
    btn_save.grid (row=2, column=6, sticky=E, padx=33, pady=10)
    buyername = str (e_buyername.get ())
    sheet['B5'] = buyername
    sheet['A4'] = 'Invoice No: '+ str(e_billno.get())
    sheet['B6'] = e_buyeradd.get('1.0', '2.0')
    sheet['B8'] = e_buyertin.get()
    sheet['H5'] = e_date.get()
    sheet['H6'] = e_vhclno.get()
    sheet['H9'] = e_place.get()
    sheet['D13'] = e_priceitem.get()
    sheet['E34'] = '=SUM(E13:E32)'
    sheet['H34'] = '=SUM(H13:H32)'
    sheet['H36'] = e_tax1amnt.get()
    sheet['H37'] = e_tax2amnt.get()
    sheet['H38'] = e_totalamnt.get()
    z = len(tree.get_children(""))
    i = 13
    j = 0
    while i<z+13:
        i += 1
        j += 1
        name = tree.item('I00%d' %j).get('text')
        values = tree.item("I00%d" %j).get('values')
        quant = float(values[0])
        unit = values[1]
        price = values[2]
        amnt = float(values[3])
        sheet['B%d' %i] = name
        sheet['D%d' % i] = price
        sheet['E%d' % i] = quant
        sheet['F%d' % i] = unit
        sheet['H%d' % i] = amnt
    excel.save ('E:\excel%d.xlsx' %n)
    global sheet1
    sheet1['A1'] = n
    pickled_in.save ('E:\pickle_n.xlsx')


btn_addbill = Button (frame_goods, text='Add', command=lambda: btn_add(flag),width=10,relief=GROOVE)
btn_addbill.grid (row=5,column=6)
y = (e_item, e_amntitem,e_priceitem)


raise_frame(main3_frame)
w.mainloop()