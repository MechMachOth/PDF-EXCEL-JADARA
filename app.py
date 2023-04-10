from tkinter import messagebox
import csv
import xlsxwriter
import tabula
from numpy import NaN
from PIL import ImageTk, Image
import sys
import os
from tkPDFViewer import tkPDFViewer as pdf
from tkinter import simpledialog, filedialog, ttk
from tkinter import *
from openpyxl import load_workbook
from fpdf import FPDF

# ==========States=======
mdp = [('')]
année = 0
mois = ''
logged = 0
traité = 0
ope = 0
bordercolor = 3
bgcolor = 0
v2 = ""
filename = ''
rje3 = 0
deb = 0
showpdf = 0
rech = 0
Dattta = []
ka = 0
reloulo = 0
kn9lb3la = ''
bach = 'att'
mselec = 0
mytag = ''
laDate = ''
changerowcolo = 0
télo = 0
bases = [("MILLIARD ", 1e9), ("MILLION ", 1e6), ("MILLE ", 1e3), ("CENT ", 1e2), ("QUATRE VINGT ", 80),
         ("SOIXANTE ", 60), ("CINQUANTE ", 50), ("QUARANTE ", 40), ("TRENTE ", 30), ("VINGT ", 20), ("DIX ", 10)]
units = ["ZERO ", "UN ", "DEUX ", "TROIS ", "QUATRE ", "CINQ ", "SIX ", "SEPT ",
         "HUIT ", "NEUF ", None, "ONZE ", "DOUZE ", "TREIZE ", "QUATORZE ", "QUINZE ", "SEIZE "]
# =======window========================
window = Tk()
icon = PhotoImage(file='logo-light.png')
window.tk.call('wm', 'iconphoto', window._w, icon)
style = ttk.Style(window)
style.theme_use("clam")

# ============Login==========

window.resizable(width=1, height=1)

checkVar1 = BooleanVar(value=False)
checkVar2 = BooleanVar(value=False)
checkVar3 = BooleanVar(value=False)
checkVar4 = BooleanVar(value=False)
checkVar5 = BooleanVar(value=False)
checkVar6 = BooleanVar(value=False)
checkVar7 = BooleanVar(value=False)
checkVar8 = BooleanVar(value=True)
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x_cordinate = int((screen_width/2) - 160)
y_cordinate = int((screen_height/2) - 140)
window.geometry("{}x{}+{}+{}".format(350, 200, x_cordinate, y_cordinate))
window.resizable(width=0, height=0)
window.resizable(width=0, height=0)
window.title("  Login")
window.config(background='black')
page_frame = Frame(window)
mdp_lb = Label(page_frame, text='Mot de passe',
               font=('Bold', 15), fg='white', bg='black')
mdp_lb.pack(pady=10)
mdp_entry = Entry(page_frame, font=('Bold', 15), bd=0, show='*')
mdp_entry.pack(pady=10)
mdp_entry.focus()


def show_passeword():
    if mdp_entry.cget('show') == '*':
        mdp_entry.config(show='')
        show_hide.config(bg='green')
    else:
        mdp_entry.config(show='*')
        show_hide.config(bg='red')


show_hide = Button(mdp_entry, bg='red', command=show_passeword)
show_hide.place(width=20, height=20, x=240, y=3)
login_btn = Button(page_frame, text='Connexion', font=(
    'Bold', 15), bd=0, bg='#158aff', fg='black', command=lambda: chek_mdp())
login_btn.pack()
error = Label(page_frame, text="", bg='black')
error.pack(pady=5)
page_frame.pack(pady=20)
page_frame.pack_propagate(False)
page_frame.configure(width=350, height=500, bg='black')
window.bind("<Return>", lambda e: chek_mdp())
window.bind("<Control-a>", lambda e: mdp_entry.select_range(0, 'end'))
mdp_entry.bind("<Escape>", lambda e: window.destroy())


def chek_mdp():
    global logged
    if mdp_entry.get() in mdp:
        if logged != 1:
            logged = 1
            error.config(text="Connexion...", bg='green')
            page_frame.after(500, lambda: page_frame.forget())
            window.after(500, lambda: acceuil())
        return
    else:
        error.config(text="try again", bg='red')
    return

# ============Acceuil==========


def acceuil():
    window.resizable(width=1, height=1)
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_cordinate = int((screen_width/2) - 160)
    y_cordinate = int((screen_height/2) - 140)
    window.geometry("{}x{}+{}+{}".format(350, 200, x_cordinate, y_cordinate))
    window.resizable(width=0, height=0)
    global rje3
    rje3 = 0
    global filename
    global Dattta
    global télo
    télo = 0
    Dattta = []
    filename = ''
    page_frame.destroy()

    def secondaire():
        zeroun='01'
        topsec = Toplevel()
        topsec.title("Attestation de don")
        topsec.geometry("500x350")
        topsec.resizable(width=0, height=0)
        icon = PhotoImage(file='logo-light.png')
        window.tk.call('wm', 'iconphoto', topsec._w, icon)
        l80 = Label(topsec, text="Date", width=20,
                    font=('Times', 11, 'bold'))
        e80 = Entry(
            topsec,  width=25)
        l80.place(x=50, y=30)
        e80.place(x=200, y=30)
        l89 = Label(topsec, text="Mois", width=20,
                    font=('Times', 11, 'bold'))
        e89 = Entry(
            topsec,  width=25)
        l89.place(x=50, y=70)
        e89.place(x=200, y=70)
        l81 = Label(topsec, text="Type", width=20,
                    font=('Times', 11, 'bold'))
        
        options = [
            "VIREMENT PERMANENT",
            "VIREMENT",
            "chèque".upper(),
            "CARTE BANCAIRE",
            "espèces".upper(),
        ]
        
        # datatype of menu text
        clicked = StringVar()
        
        # initial menu text
        clicked.set( "VIREMENT" )
        
        # Create Dropdown menu
        drop = OptionMenu( topsec , clicked , *options )
        drop.place(x=200, y=105)
        l81.place(x=50, y=110)
        l82 = Label(topsec, text="Nom donneur d'ordre",
                    width=20, font=('Times', 11, 'bold'))
        e82 = Entry(
            topsec, width=25)
        l82.place(x=50, y=150)
        e82.place(x=200, y=150)
        l83 = Label(topsec, text="Montant",
                    width=20, font=('Times', 11, 'bold'))
        e83 = Entry(
            topsec, width=25)
        l83.place(x=50, y=190)
        e83.place(x=200, y=190)
        l85 = Label(topsec, text="Montant en lettre",
                    width=20, font=('Times', 11, 'bold'))
        e85 = Entry(
            topsec,  width=25)
        l85.place(x=50, y=230)
        e85.place(x=200, y=230)

        l86 = Label(topsec, text="N° Attestation",
                    width=20, font=('Times', 11, 'bold'))
        l87 = Label(topsec, text="D-AA-MM-'",
                    width=20, font=('Times', 11, 'bold'))
        e86 = Entry(
            topsec, width=10)
        l86.place(x=50, y=270)
        l87.place(x=170, y=275)
        e86.place(x=300, y=270)
        e80.focus()
        e89.configure(state="disabled")
        e85.configure(state="disabled")

        def idcode(e):
            global Dattta
            l87.configure(
                text='D-'+e80.get()[6: len(e80.get())]+'-'+e80.get()[3:5]+'-')
            if e80.get() == '':
                l87.configure(text='D-AA-MM-')
        e80.bind("<KeyRelease>", idcode)
        def hantaTchouf():
            e80.delete(0, END)
            e82.delete(0, END)
            e83.delete(0, END)
            e85.delete(0, END)
            e86.delete(0, END)
            topsec.destroy()

        def attsec():
            attoto=0
            Fait_le = e80.get().upper()
            heywdi=e86.get().upper()
            if heywdi=='1' or heywdi=='':
                heywdi='01'
            if heywdi=='2':
                heywdi='02'
            if heywdi=='3':
                heywdi='03'
            if heywdi=='4':
                heywdi='04'
            if heywdi=='5':
                heywdi='05'
            if heywdi=='6':
                heywdi='06'
            if heywdi=='7':
                heywdi='07'
            if heywdi=='8':
                heywdi='08'
            if heywdi=='9':
                heywdi='09'
            D_22_09_10 = "D"+"-"+Fait_le[3:5]+"-"+Fait_le[6:len(Fait_le)]+"-"+heywdi
            Montant = e83.get().upper()
            Montant_en_lettre = e85.get().upper()
            Nom = e82.get().upper()
            Type = clicked.get().upper()
            mois=''
            kok=0
            if os.path.exists(str(D_22_09_10)+".pdf"): 
                attoto=1
            def chiftolett(value, skip=-1):
                if int(value) < len(units) and units[int(value)]:
                    return [] if int(value) <= skip else [units[int(value)]]
                for name, v in bases:
                    if int(value) >= v:
                        return chiftolett(int(int(value)/v), 1 if v <= 1000 else -1) + [name] + chiftolett(int(int(value) % v), 0)
            
            if Montant == '' or Fait_le == '':
                messagebox.showinfo(
                    title='Erreur !!', message="la date et le montant sont obligatoire.", parent=topsec)
                ope = 0
                return
            
            if Montant.count(',') == 1:
                if len(Montant)-Montant.index(',') != 3:
                    messagebox.showinfo(
                        title='Erreur !!', message="Montant invalide, merci de saisir deux chiffres après la virgule.", parent=topsec)
                    return
                else:
                    for i in Montant:
                        if i == ',' or i == '.':
                            pass
                        elif not i.isnumeric():
                            messagebox.showinfo(
                                title='Erreur !!', message="Montant invalide", parent=topsec)
                            return
                    po = Montant[len(Montant)-2] + \
                        Montant[len(Montant)-1]
                    if po == '00':
                        Montant_en_lettre = "".join(chiftolett(
                            int(int(Montant.replace(',', '').replace('.', ''))/100)))
                    else:
                        Montant_en_lettre = "".join(chiftolett(
                            int(int(Montant.replace(',', '').replace('.', ''))/100)))
                        Montant_en_lettre += 'VIRGULE '
                        if po == "01" or po == "02" or po == "03" or po == "04" or po == "05" or po == "06" or po == "07" or po == "08" or po == "09":
                            Montant_en_lettre += "ZERO "
                        Montant_en_lettre += "".join(chiftolett(int(po)))
                x=''
                Montant=Montant.replace('.','')
                if len(Montant)>4 :
                    bima=Montant[len(Montant)-3: len(Montant)]
                    Montant=Montant[0:len(Montant)-3]
                    for i in range(len(Montant)-1,-1,-1) : 
                        if kok==3 or kok==6 or kok==9 or kok==12 or kok==15 or kok==18 or kok==21:
                            x=x+'.'
                            x= x+Montant[i]
                        else : 
                            x= x+Montant[i]
                        kok=kok+1
                if kok!=0 : 
                    Montant=''
                    for i in range(len(x)-1,-1,-1):
                        Montant=Montant+x[i]
                    Montant=Montant+bima
            else:
                x=''
                Montant=Montant.replace('.','')
                if len(Montant)>4 :
                    for i in range(len(Montant)-1,-1,-1) : 
                        if kok==3 or kok==6 or kok==9 or kok==12 or kok==15 or kok==18 or kok==21:
                            x=x+'.'
                            x= x+Montant[i]
                        else : 
                            x= x+Montant[i]
                        kok=kok+1
                if kok!=0 : 
                    Montant=''
                    for i in range(len(x)-1,-1,-1):
                        Montant=Montant+x[i]
                for i in Montant:
                    if i == '.':
                        pass
                    elif not i.isnumeric():
                        messagebox.showinfo(
                            title='Erreur !!', message="Montant invalide", parent=topsec)
                        return
                Montant_en_lettre = "".join(chiftolett(
                    int(int(Montant.replace(',', '').replace('.', '')))))
                Montant = "".join((Montant, ',00'))

            if Fait_le[0:2].isnumeric() and Fait_le[3:5].isnumeric() and Fait_le[6:8].isnumeric() and len(Fait_le) == 8:
                if int(Fait_le[3:5]) > 12 or int(Fait_le[0:2]) > 31:
                    messagebox.showinfo(
                        title='Erreur !!', message="* Le nombre de jour ne doit pas être superieur à 31. \n\n* Le nombre de mois ne doit pas être superieur à 12.", parent=topsec)
                    return
            else:
                messagebox.showinfo(
                    title='Erreur !!', message="La date doit être sous la forme suivante :\n\n            'JJxMMxAA'", parent=topsec)
                return
            hoho = list(Fait_le)
            hoho[2] = '/'
            hoho[5] = '/'
            Fait_le = ''.join(hoho)
            if Fait_le[3:5] == '01':
                mois = 'Janvier'
            if Fait_le[3:5] == '02':
                mois = 'Février'
            if Fait_le[3:5] == '03':
                mois = 'Mars'
            if Fait_le[3:5] == '04':
                mois = 'Avril'
            if Fait_le[3:5] == '05':
                mois = 'Mai'
            if Fait_le[3:5] == '06':
                mois = 'Juin'
            if Fait_le[3:5] == '07':
                mois = 'Juillet'
            if Fait_le[3:5] == '08':
                mois = 'Août'
            if Fait_le[3:5] == '09':
                mois = 'Septembre'
            if Fait_le[3:5] == '10':
                mois = 'Octobre'
            if Fait_le[3:5] == '11':
                mois = 'Novembre'
            if Fait_le[3:5] == '12':
                mois = 'Décembre'
            pdf = FPDF(orientation='P', format='A4')
            pdf.add_page()

            pdf.set_xy(90,65)
            pdf.set_font("times",size=21,style='BU')
            pdf.cell(txt='Attestation de Don',w=30,align='C')
            pdf.image('logo-att.png',80, 10,w = 50,h=40)
            pdf.set_xy(91.5,75)
            pdf.cell(txt=D_22_09_10,w=28,align='C')

            text1="Je soussignée, Mme Bouchra OUTAGHANI, Trésorière Générale de JADARA FOUNDATION, atteste par la présente avoir reçu la somme de "
            text2=Montant+" dirhams ( "+Montant_en_lettre+"dirhams )"
            if Type.lower()=="espèce" or Type.lower()=="espece":
                Type="espèce".upper()
                text3=" en "
            else:
                text3=" par "
            text4=Type.upper()+" "
            text5="de "
            text6=Nom.upper()+"."
            text7="La contribution de "
            text8=Nom+" "
            text9="participera au financement de la mission de JADARA Foudation telle que prévue par ses status accessibles sur son site web www.jadara.foundation."
            text10=Nom.upper()+" "
            text11="peut accéder au rapport morale et financier de JADARA Foudation sur son site web www.jadara.foudation."
            text12="Jadara Foudation est une association reconnue d'utilité publique par le Décret N°2-13-339 du 29 Avril 2013 tel que modifié par le Décret N°2.22.834 du 19 octobre 2022."
            text13="Cette attestation est délivrée pour servir et valoir ce que de droit."

            text14="Fait à Casablanca, le "
            text15=Fait_le[0:2]+" "+mois+" "+"20"+Fait_le[6:len(Fait_le)]
            text16="Bouchra OUTAGHANI"

            pdf.set_auto_page_break("ON", margin = 0.0)
            pdf.set_font("times",size=12)
            pdf.set_xy(20,105)
            pdf.multi_cell(w=170,h=5,txt=text1+"**"+text2+"**"+text3+"**"+text4+"**"+text5+"**"+text6+"**"+"\n\n"+text7+"**"+text8+"**"+text9+"\n\n"+"**"+text10+"**"+text11+"\n\n"+text12+"\n\n"+text13,markdown=True,
                            align='L')

            pdf.set_font("times",size=11)
            pdf.set_xy(100,200)
            pdf.multi_cell(w=90,h=5,txt=text14+"**"+text15+"**"+"\n\n"+"**"+text16+"**",markdown=True,align='R')
            pdf.set_xy(100,215)
            pdf.multi_cell(w=90,h=5,txt="**Trésorière Générale**",markdown=True,align='R')
            pdf.set_font("times",size=9)
            pdf.set_xy(100,220)
            pdf.multi_cell(w=90,h=5,txt="**P.O**",markdown=True,align='R')
            pdf.set_xy(100,225)
            pdf.multi_cell(w=90,h=5,txt="**Bochra CHABBOUBA ELIDRISSI**",markdown=True,align='R')
            pdf.set_xy(100,230)
            pdf.multi_cell(w=90,h=5,txt="**Responsable Administrative et Financière**",markdown=True,align='R')



            pdf.set_fill_color(193, 153, 9)
            pdf.set_xy(8,275)
            pdf.multi_cell(w=0,h=0.5,txt="",fill=True)

            pdf.set_text_color(45, 82, 158)
            pdf.set_font("times",size=14,style="B")
            pdf.set_xy(8,280)
            pdf.multi_cell(w=0,h=5,txt="JADARA Foundation")

            pdf.set_text_color(193, 153, 9)
            pdf.set_font("times",size=7.5,style="")
            pdf.set_xy(8,285)
            pdf.multi_cell(w=0,h=5,txt="Jadara Foundation, association reconnue d'utilité publique selon de décret N°2.22.834")

            pdf.set_text_color(0, 0, 0)
            pdf.set_font("times",size=7.5,style="")
            pdf.set_xy(8,289)
            pdf.multi_cell(w=0,h=5,txt="**Adresse**: 295, Bd Abdelmoumen C24, angle Rue la Percée, 20100, Casablanca",markdown=True)


            pdf.set_font("times",size=8,style="")
            pdf.set_xy(107,279)     
            pdf.multi_cell(w=40,h=5,txt="**T**: +212 522 861 880\n**F**: +212 522 864 178\n**Mail**: contact@jadara.foundation",markdown=True)

            pdf.set_font("times",size=8,style="")
            pdf.set_xy(158,283)     
            pdf.multi_cell(w=40,h=5,txt="**www.jadara.foundation**",markdown=True)


            pdf.set_xy(152,275)
            pdf.multi_cell(w=0.5,h=20,txt="",fill=True)


            pdf.set_xy(102,275)
            pdf.multi_cell(w=0.5,h=20,txt="",fill=True)




            pdf.output(str(D_22_09_10)+".pdf")
            if attoto==0:
                messagebox.showinfo(
                    title='', message="Fichier "+str(D_22_09_10)+".pdf créé. ", parent=topsec)
            else : 
                messagebox.showinfo(
                    title='', message="Fichier "+str(D_22_09_10)+".pdf mis à jour. ", parent=topsec)
            hantaTchouf()

        submitbutton = Button(
            topsec, text='Enregistrer', command=lambda: attsec())
        submitbutton.configure(
            font=('Times', 11, 'bold'), bg='green', fg='white')
        submitbutton.place(x=300, y=310)
        cancelbutton = Button(
            topsec, text='Annulé', command=lambda: hantaTchouf())
        cancelbutton.configure(
            font=('Times', 11, 'bold'), bg='red', fg='white')
        cancelbutton.place(x=200, y=310)
        topsec.protocol(
            "WM_DELETE_WINDOW", hantaTchouf)
        topsec.bind(
            "<Return>", lambda e: attsec())
        topsec.bind(
            "<Escape>", lambda e: hantaTchouf())
        return

    def principal():
        global rje3
        global v2
        global v1
        global filename
        global bordercolor
        global bgcolor
        rje3 = 1
        b.place_forget()
        a.place_forget()
        lo13.place_forget()
        lo14.place_forget()
        window.resizable(width=1, height=1)
        window.geometry("5000x5000")

        def reset():
            sur = messagebox.askquestion(
                'les données seront perdues !', "êtes-vous sûr de vouloir relancer l'application ?", icon='warning')
            if sur == 'yes':
                python = sys.executable
                os.execl(python, python, * sys.argv)
        frame2 = Frame(window)
        frame1 = Frame(window)
        frame4 = Frame(frame1)
        frame3 = Frame(frame1)
        frame5 = Frame(frame1, height=350, width=45, bg='black')
        frame4.pack(padx=0, pady=5)
        frame5.place(x=5, y=202)
        thblack = Button(frame4, text='',
                         command=lambda: borco())
        thblack.pack(pady=5, padx=30, side='left')
        thblackbg = Button(frame4, text='',
                           command=lambda: bgco())
        thblackbg.pack(pady=5, padx=30, side='left')

        if bordercolor == 2:
            frame1.configure(bg='black',
                             highlightbackground="yellow", highlightthickness=2)
            frame2.configure(bg='black',
                             highlightbackground="yellow", highlightthickness=2)
            frame3.configure(bg='black',
                             highlightbackground="yellow", highlightthickness=2)
            frame4.configure(bg='black',
                             highlightbackground="yellow", highlightthickness=2)
            thblack.configure(background='#000080')

        elif bordercolor == 1:
            frame4.configure(bg='black',
                             highlightbackground="red", highlightthickness=2)
            frame3.configure(bg='black',
                             highlightbackground="red", highlightthickness=2)
            frame2.configure(bg='black',
                             highlightbackground="red", highlightthickness=2)
            frame1.configure(bg='black',
                             highlightbackground="red", highlightthickness=2)
            thblack.configure(background='yellow')

        elif bordercolor == 3:

            frame1.configure(bg='black',
                             highlightbackground="#000080", highlightthickness=2)
            frame2.configure(bg='black',
                             highlightbackground="#000080", highlightthickness=2)
            frame3.configure(bg='black',
                             highlightbackground="#000080", highlightthickness=2)
            frame4.configure(bg='black',
                             highlightbackground="#000080", highlightthickness=2)
            thblack.configure(background='black')

        elif bordercolor == 0:
            frame4.configure(bg='black',
                             highlightbackground="green", highlightthickness=2)
            frame3.configure(bg='black',
                             highlightbackground="green", highlightthickness=2)
            frame2.configure(bg='black',
                             highlightbackground="green", highlightthickness=2)
            frame1.configure(bg='black',
                             highlightbackground="green", highlightthickness=2)
            thblack.configure(background='red')

        elif bordercolor == 5:
            frame1.configure(bg='black',
                             highlightbackground="white", highlightthickness=2)
            frame2.configure(bg='black',
                             highlightbackground="white", highlightthickness=2)
            frame3.configure(bg='black',
                             highlightbackground="white", highlightthickness=2)
            frame4.configure(bg='black',
                             highlightbackground="white", highlightthickness=2)
            thblack.configure(background='green')

        elif bordercolor == 4:
            frame4.configure(bg='black',
                             highlightbackground="black", highlightthickness=2)
            frame3.configure(bg='black',
                             highlightbackground="black", highlightthickness=2)
            frame2.configure(bg='black',
                             highlightbackground="black", highlightthickness=2)
            frame1.configure(bg='black',
                             highlightbackground="black", highlightthickness=2)
            thblack.configure(background='white')

        def bgco():
            global bgcolor
            if bgcolor == 0:
                bgcolor = 1
                window.configure(background='white')
                frame1.configure(bg='white')
                frame2.configure(bg='white')
                frame3.configure(bg='white')
                frame4.configure(bg='white')
                frame5.configure(bg='white')
                thblackbg.configure(bg='green')
            elif bgcolor == 1:
                bgcolor = 2
                window.configure(background='green')
                frame1.configure(bg='green')
                frame2.configure(bg='green')
                frame5.configure(bg='green')
                frame3.configure(bg='green')
                thblackbg.configure(bg='yellow')
                frame4.configure(bg='green')
            elif bgcolor == 2:
                bgcolor = 3
                window.configure(background='yellow')
                frame1.configure(bg='yellow')
                frame2.configure(bg='yellow')
                frame3.configure(bg='yellow')
                frame5.configure(bg='yellow')
                thblackbg.configure(bg='#000080')
                frame4.configure(bg='yellow')
            elif bgcolor == 3:
                bgcolor = 4
                window.configure(background='#000080')
                frame1.configure(bg='#000080')
                frame5.configure(bg='#000080')
                frame2.configure(bg='#000080')
                frame3.configure(bg='#000080')
                thblackbg.configure(bg='red')
                frame4.configure(bg='#000080')
            elif bgcolor == 4:
                bgcolor = 5
                window.configure(background='red')
                frame1.configure(bg='red')
                frame2.configure(bg='red')
                frame5.configure(bg='red')
                frame3.configure(bg='red')
                thblackbg.configure(bg='black')
                frame4.configure(bg='red')
            elif bgcolor == 5:
                bgcolor = 0
                window.configure(background='black')
                frame1.configure(bg='black')
                frame2.configure(bg='black')
                frame3.configure(bg='black')
                frame5.configure(bg='black')
                frame4.configure(bg='black')
                thblackbg.configure(bg='white')

            return

        if bgcolor == 0:
            window.configure(background='black')
            frame1.configure(bg='black')
            frame2.configure(bg='black')
            frame3.configure(bg='black')
            frame5.configure(bg='black')
            frame4.configure(bg='black')
            thblackbg.configure(background='white')

        if bgcolor == 1:
            window.configure(background='white')
            frame1.configure(bg='white')
            frame2.configure(bg='white')
            frame3.configure(bg='white')
            frame5.configure(bg='white')
            frame4.configure(bg='white')
            thblackbg.configure(background='green')

        if bgcolor == 2:
            window.configure(background='green')
            frame1.configure(bg='green')
            frame2.configure(bg='green')
            frame3.configure(bg='green')
            frame5.configure(bg='green')
            frame4.configure(bg='green')
            thblackbg.configure(background='yellow')

        if bgcolor == 3:
            window.configure(background='yellow')
            frame1.configure(bg='yellow')
            frame2.configure(bg='yellow')
            frame3.configure(bg='yellow')
            frame5.configure(bg='yellow')
            frame4.configure(bg='yellow')
            thblackbg.configure(background='#000080')

        if bgcolor == 4:
            window.configure(background='#000080')
            frame1.configure(bg='#000080')
            frame2.configure(bg='#000080')
            frame3.configure(bg='#000080')
            frame5.configure(bg='#000080')
            frame4.configure(bg='#000080')
            thblackbg.configure(background='red')

        if bgcolor == 5:
            window.configure(background='red')
            frame1.configure(bg='red')
            frame2.configure(bg='red')
            frame3.configure(bg='red')
            frame5.configure(bg='red')
            frame4.configure(bg='red')
            thblackbg.configure(background='black')

        def borco():
            global bordercolor
            if bordercolor == 0:
                frame1.configure(
                    highlightbackground="red", highlightthickness=2)
                frame2.configure(
                    highlightbackground="red", highlightthickness=2)
                frame3.configure(
                    highlightbackground="red", highlightthickness=2)
                frame4.configure(
                    highlightbackground="red", highlightthickness=2)
                thblack.configure(background='yellow')
                bordercolor = 1
            elif bordercolor == 2:
                bordercolor = 3
                frame1.configure(
                    highlightbackground="#000080", highlightthickness=2)
                frame2.configure(
                    highlightbackground="#000080", highlightthickness=2)
                frame4.configure(
                    highlightbackground="#000080", highlightthickness=2)
                frame3.configure(
                    highlightbackground="#000080", highlightthickness=2)
                thblack.configure(background='black')
            elif bordercolor == 1:
                bordercolor = 2
                frame1.configure(
                    highlightbackground="yellow", highlightthickness=2)
                frame2.configure(
                    highlightbackground="yellow", highlightthickness=2)
                frame3.configure(
                    highlightbackground="yellow", highlightthickness=2)
                frame4.configure(
                    highlightbackground="yellow", highlightthickness=2)
                thblack.configure(background='#000080')
            elif bordercolor == 3:
                bordercolor = 4
                frame1.configure(
                    highlightbackground="black", highlightthickness=2)
                frame2.configure(
                    highlightbackground="black", highlightthickness=2)
                frame3.configure(
                    highlightbackground="black", highlightthickness=2)
                frame4.configure(
                    highlightbackground="black", highlightthickness=2)
                thblack.configure(background='white')
            elif bordercolor == 4:
                bordercolor = 5
                frame1.configure(
                    highlightbackground="white", highlightthickness=2)
                frame2.configure(
                    highlightbackground="white", highlightthickness=2)
                frame3.configure(
                    highlightbackground="white", highlightthickness=2)
                frame4.configure(
                    highlightbackground="white", highlightthickness=2)
                thblack.configure(background='green')
            elif bordercolor == 5:
                bordercolor = 0
                frame1.configure(
                    highlightbackground="green", highlightthickness=2)
                frame2.configure(
                    highlightbackground="green", highlightthickness=2)
                frame3.configure(
                    highlightbackground="green", highlightthickness=2)
                frame4.configure(
                    highlightbackground="green", highlightthickness=2)
                thblack.configure(background='red')

        # ==== PDFViewer=====
        filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title='Select pdf file',
                                              filetypes=(("PDF File", ".pdf"), ("PDF File", ".PDF"), ("All file", ".txt")))
        if type(filename) == tuple or filename == '':
            acceuil()
            return
        v1 = pdf.ShowPdf()
        v2 = v1.pdf_view(frame2, pdf_location=open(filename, 'r'), width=75)
        v2.pack(fill='y', side='left', expand=True)
        b.place_forget()
        a.place_forget()

        def rj3():
            global traité
            if rje3 == 1:
                sur = messagebox.askquestion(
                    '!!!', "les données seront perdues !", icon='warning')
                if sur == 'yes':
                    btn_re.place_forget()
                    frame1.destroy()
                    frame2.destroy()
                    frame3.destroy()
                    traité = 0
                    lo30.place_forget()
                    acceuil()
        window.bind("<Escape>", lambda e: rj3())
        btn_re = Button(window, text='<<--', command=lambda: rj3())
        btn_re.place(x=0, y=0)
        lo30 = Label(text='Echap', fg='white',
                     background='black', font=('Times', 10))

        def traitement():
            global traité
            global Dattta
            global laDate
            charger.config(bg='gray')
            if traité == 0:
                Dattta = []
                global année
                charger.config(text='Chargement...')
                année = simpledialog.askinteger(
                    "Année", "Merci de Saisir l'année", initialvalue=2023, parent=None, maxvalue=9999, minvalue=1000)
                traité = 1
                lo30.place(x=10, y=35)
                if année == None:
                    charger.config(text='Charger les données')
                    traité = 0
                    return
                charger.pack_forget()
                lo20.pack_forget()
                v2.forget()
                tables = tabula.read_pdf(filename, pages='all')
                table = []
                for i in range(len(tables)):
                    tables[i].to_csv(f"table{i}.csv")
                for j in range(len(tables)):
                    with open(f"table{j}.csv", 'r') as file:
                        reader = csv.reader(file)
                        for each in reader:
                            if (each[1] != ''):
                                if (each[len(each)-1] != ''):
                                    if each.count("Unnamed: 2") == 0:
                                        if each[4] == '':
                                            if each.count("Page N°") == 0:
                                                for i in each:
                                                    if i == '':
                                                        each.remove(i)
                                                        for i in each:
                                                            if i == '':
                                                                each.remove(i)
                                                table.append(each)
                year = année-2000
                i = 1
                for each in table:
                    each[1] = each[1][0:5]+'/'+str(year)
                    each[0] = 'D-' + str(year) + '-'+each[1][3:5]+'-'+str(i)
                    i = i+1
                    type = each[2][0:4]
                    if type == 'VIRE':
                        if each[2].count('DE') == 2:
                            each.insert(3, each[2][29:])
                        if each[2].count('DE') < 2:
                            each.insert(3, each[2][24:])
                        each[2] = 'VIREMENT PERMANENT'
                    if type == 'VIR ':
                        each.insert(3, each[2][26:])
                        each[2] = 'VIREMENT'
                    if type == 'VERS':
                        each.insert(3, each[2][13:])
                        each[2] = 'VERSEMENT'
                    if type == 'CRED':
                        each.insert(3, "")
                    if type == 'ENCA':
                        each.insert(3, "")
                    if type == 'OBJE':
                        each.insert(3, "")
                table.pop(len(table)-1)
                for j in range(len(tables)):
                    os.remove(f"table{j}.csv")

                style = ttk.Style()
                style.theme_use('default')
                style.configure("Treeview", foreground="black",
                                fieldbackground="silver", rowheight=25)
                scrolly = ttk.Scrollbar(frame2, orient=VERTICAL)
                my_tree = ttk.Treeview(
                    frame2, height=37, yscrollcommand=scrolly.set)
                my_tree.tag_configure('gray', background='gray')
                my_tree.tag_configure('normal', background='white')
                my_tree.tag_configure('blue', background='lightblue')
                my_tree.tag_configure('green', background='lightgreen')
                my_tree.tag_configure('red', background='red')
                my_tree['columns'] = ("Date", "Mois", "Type", "NOM Donneur d'ordre",
                                      "Montant", "Détail", "Montant en lettre", "N° Attestation")
                my_tree.column("#0", width=0, stretch=NO)
                my_tree.column("Date", width=80, anchor=CENTER, minwidth=25)
                my_tree.column("Mois", width=90, anchor=CENTER, minwidth=25)
                my_tree.column("Type", width=280, anchor=CENTER, minwidth=25)
                my_tree.column("NOM Donneur d'ordre", width=395,
                               anchor=CENTER, minwidth=25)
                my_tree.column("Montant", width=170,
                               anchor=CENTER, minwidth=25)
                my_tree.column("Détail", width=80, anchor=CENTER, minwidth=25)
                my_tree.column("Montant en lettre", width=405,
                               anchor=CENTER, minwidth=25)
                my_tree.column("N° Attestation", width=120,
                               anchor=CENTER, minwidth=25)
                my_tree.heading("#0", text="", anchor=W)
                my_tree.heading("Date", text="Date", anchor=CENTER)
                my_tree.heading("Mois", text="Mois", anchor=CENTER)
                my_tree.heading("Type", text="Type", anchor=CENTER)
                my_tree.heading("NOM Donneur d'ordre",
                                text="NOM Donneur d'ordre", anchor=CENTER)
                my_tree.heading("Montant", text="Montant", anchor=CENTER)
                my_tree.heading("Détail", text="Détail", anchor=CENTER)
                my_tree.heading("Montant en lettre",
                                text="Montant en lettre", anchor=CENTER)
                my_tree.heading("N° Attestation",
                                text="N° Attestation", anchor=CENTER)
                my_tree.pack(padx=5, fill='y')
                scrolly.configure(command=my_tree.yview)
                scrolly.place(y=70, height=860, x=1655)
                laDate = table[1][1]
                for i in table:
                    if i[1][3:5] == '01':
                        mois = 'JANVIER'
                    if i[1][3:5] == '02':
                        mois = 'FEVRIER'
                    if i[1][3:5] == '03':
                        mois = 'MARS'
                    if i[1][3:5] == '04':
                        mois = 'AVRIL'
                    if i[1][3:5] == '05':
                        mois = 'MAI'
                    if i[1][3:5] == '06':
                        mois = 'JUIN'
                    if i[1][3:5] == '07':
                        mois = 'JUILLET'
                    if i[1][3:5] == '08':
                        mois = 'AOUT'
                    if i[1][3:5] == '09':
                        mois = 'SEPTEMBRE'
                    if i[1][3:5] == '10':
                        mois = 'OCTOBRE'
                    if i[1][3:5] == '11':
                        mois = 'NOVEMBRE'
                    if i[1][3:5] == '12':
                        mois = 'DECEMBRE'

                    def chiftolett(value, skip=-1):
                        if int(value) < len(units) and units[int(value)]:
                            return [] if int(value) <= skip else [units[int(value)]]
                        for name, v in bases:
                            if int(value) >= v:
                                return chiftolett(int(int(value)/v), 1 if v <= 1000 else -1) + [name] + chiftolett(int(int(value) % v), 0)
                    lenght = len(i[4])
                    po = i[4][lenght-2]+i[4][lenght-1]
                    l = len(i[2])
                    mytag = 'normal'
                    if i[2][0:6] == 'ENCAIS':
                        i[2]='chèque'.upper()
                    if i[2][0:l] == 'CREDIT COMMERCANT':
                        mytag = 'blue'
                    if i[2][0:l] != 'CREDIT COMMERCANT' and i[3] == '':
                        mytag = 'red'
                    if mytag == 'blue':
                        if po == '00':
                            my_tree.insert(parent='', index='end', iid=i, text='', values=(i[1], mois, i[2].replace('CREDIT COMMERCANT', 'CARTE BANCAIRE'), i[3], i[4], "", "".join(
                                chiftolett(int(int(i[4].replace(',', '').replace('.', ''))/100))), i[0]), tags=(mytag))
                            mytag = 'normal'
                        else:
                            ntl = "".join(chiftolett(
                                int(int(i[4].replace(',', '').replace('.', ''))/100)))
                            ntl += 'VIRGULE '
                            ntl += "".join(chiftolett(int(po)))
                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                i[1], mois, i[2].replace('CREDIT COMMERCANT', 'CARTE BANCAIRE'), i[3], i[4], "", ntl, i[0]), tags=(mytag))
                            mytag = 'normal'
                    else:
                        if po == '00':
                            my_tree.insert(parent='', index='end', iid=i, text='', values=(i[1], mois, i[2].replace('CREDIT COMMERCANT', 'CARTE BANCAIRE').replace('ENCAISSEMENT CHEQUE NUM - ', ''), i[3], i[4], "", "".join(
                                chiftolett(int(int(i[4].replace(',', '').replace('.', ''))/100))), i[0]), tags=(mytag))
                            mytag = 'normal'
                        else:
                            ntl = "".join(chiftolett(
                                int(int(i[4].replace(',', '').replace('.', ''))/100)))
                            ntl += 'VIRGULE '
                            ntl += "".join(chiftolett(int(po)))
                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                i[1], mois, i[2].replace('CREDIT COMMERCANT', 'CARTE BANCAIRE'), i[3], i[4], "", ntl, i[0]), tags=(mytag))
                            mytag = 'normal'

                children = my_tree.get_children()
                for parent in my_tree.get_children():
                    Dattta.append(my_tree.item(parent)["values"])

                pdff = Button(frame1, text='PDF',
                              command=lambda: showPDF())
                pdff.pack(padx=20)
                lo = Label(frame1, text='Ctrl-P', fg='white',
                           background='black', font=('Times', 10))
                lo.pack()
                frame3.pack(fill='y', padx=10, pady=10)
                Date = StringVar()
                Mois = StringVar()
                Type = StringVar()
                Nom = StringVar()
                Montant = StringVar()
                détail = StringVar()
                Montant_en_lettre = StringVar()
                NAttestation = StringVar()

                def showPDF():
                    global v2
                    global showpdf
                    if showpdf == 0:
                        my_tree.pack_forget()
                        v2.pack(fill='y', side='left', expand=True)
                        pdff.configure(text='DATA')
                        showpdf = 1
                        scrolly.place_forget()
                        frame5.place_forget()
                        delete.pack_forget()
                        edite.pack_forget()
                        Attestation.pack_forget()
                        lo2.pack_forget()
                        lo4.pack_forget()
                        lo5.pack_forget()
                        rowscolor.place_forget()

                    else:
                        add.pack_forget()
                        lo3.pack_forget()
                        pdff.configure(text='PDF')
                        showpdf = 0
                        v2.pack_forget()
                        my_tree.pack()
                        scrolly.place(y=70, height=860, x=1655)
                        frame5.place(x=5, y=202)
                        delete.pack(pady=20, padx=20)
                        lo2.pack()
                        add.pack(pady=20, padx=20)
                        lo3.pack()
                        edite.pack(pady=20, padx=20)
                        lo4.pack()
                        Attestation.pack(pady=20, padx=20)
                        lo5.pack()
                        rowscolor.place(x=75, y=5)

                def editData(tree):
                    global ope
                    if ope == 0:
                        ope = 1
                        curItem = tree.focus()
                        values = tree.item(curItem, "values")
                        if len(tree.selection()) < 2:
                            if len(curItem) > 0:
                                topEdit = Toplevel()
                                topEdit.title("Modifier")
                                topEdit.geometry("500x400")
                                icon = PhotoImage(file='logo-light.png')
                                window.tk.call(
                                    'wm', 'iconphoto', topEdit._w, icon)
                                topEdit.resizable(width=0, height=0)
                                l9 = Label(topEdit, text="Date", width=20,
                                           font=('Times', 11, 'bold'))
                                e9 = Entry(
                                    topEdit, textvariable=Date, width=25)
                                l9.place(x=50, y=30)
                                e9.place(x=200, y=30)

                                l10 = Label(topEdit, text="Mois", width=20,
                                            font=('Times', 11, 'bold'))
                                e10 = Entry(
                                    topEdit, textvariable=Mois, width=25)
                                l10.place(x=50, y=70)
                                e10.place(x=200, y=70)

                                l11 = Label(topEdit, text="Type", width=20,
                                            font=('Times', 11, 'bold'))
                                
                                options = [
                                    "VIREMENT PERMANENT",
                                    "VIREMENT",
                                    "chèque".upper(),
                                    "CARTE BANCAIRE",
                                    "espèces".upper(),
                                ]
                                
                                # datatype of menu text
                                clicked = StringVar()
                                
                                # initial menu text
                                clicked.set( values[2] )
                                
                                # Create Dropdown menu
                                drop = OptionMenu( topEdit , clicked , *options )
                                drop.place(x=200, y=105)
                                l11.place(x=50, y=110)

                                l12 = Label(topEdit, text="Nom donneur d'ordre",
                                            width=20, font=('Times', 11, 'bold'))
                                e12 = Entry(
                                    topEdit, textvariable=Nom, width=25)
                                l12.place(x=50, y=150)
                                e12.place(x=200, y=150)

                                l13 = Label(topEdit, text="Montant",
                                            width=20, font=('Times', 11, 'bold'))
                                e13 = Entry(
                                    topEdit, textvariable=Montant, width=25)
                                l13.place(x=50, y=190)
                                e13.place(x=200, y=190)

                                l14 = Label(topEdit, text="Détail",
                                            width=20, font=('Times', 11, 'bold'))
                                e14 = Entry(
                                    topEdit, textvariable=détail, width=25)
                                l14.place(x=50, y=230)
                                e14.place(x=200, y=230)

                                l15 = Label(topEdit, text="Montant en lettre",
                                            width=20, font=('Times', 11, 'bold'))
                                e15 = Entry(
                                    topEdit, textvariable=Montant_en_lettre, width=25)
                                l15.place(x=50, y=270)
                                e15.place(x=200, y=270)

                                l16 = Label(topEdit, text="N° Attestation",
                                            width=20, font=('Times', 11, 'bold'))
                                e16 = Entry(
                                    topEdit, textvariable=NAttestation, width=25)
                                l16.place(x=50, y=310)
                                e16.place(x=200, y=310)
                                e9.focus()

                                def insertData(tree):
                                    global ope
                                    nonlocal e9, e10, e12, e13, e14, e15, e16
                                    global Dattta
                                    da = Date.get().strip()
                                    mo = Mois.get().strip()
                                    ty = clicked.get().upper()
                                    no = Nom.get().strip().upper()
                                    mon = Montant.get().strip()
                                    dé = détail.get().strip()
                                    mol = Montant_en_lettre.get().strip()
                                    att = NAttestation.get().strip()
                                    x=""
                                    kok=0

                                    if mon.count(',') == 1:
                                        if len(mon)-mon.index(',') != 3:
                                            messagebox.showinfo(
                                                title='Erreur !!', message="Montant invalide, merci de saisir deux chiffres après la virgule.", parent=topEdit)
                                            ope = 0
                                            return
                                        else:
                                            for i in mon:
                                                if i == ',' or i == '.':
                                                    pass
                                                elif not i.isnumeric():
                                                    messagebox.showinfo(
                                                        title='Erreur !!', message="Montant invalide", parent=topEdit)
                                                    ope = 0
                                                    return
                                            po = mon[len(mon)-2] + \
                                                mon[len(mon)-1]
                                            if po == '00':
                                                mol = "".join(chiftolett(
                                                    int(int(mon.replace(',', '').replace('.', ''))/100)))
                                            else:
                                                mol = "".join(chiftolett(
                                                    int(int(mon.replace(',', '').replace('.', ''))/100)))
                                                mol += 'VIRGULE '
                                                if po == "01" or po == "02" or po == "03" or po == "04" or po == "05" or po == "06" or po == "07" or po == "08" or po == "09":
                                                    mol += "ZERO "
                                                mol += "".join(chiftolett(int(po)))
                                    else:
                                        for i in mon:
                                            if i == '.':
                                                pass
                                            elif not i.isnumeric():
                                                messagebox.showinfo(
                                                    title='Erreur !!', message="Montant invalide", parent=topEdit)
                                                ope = 0
                                                return
                                        mol = "".join(chiftolett(
                                            int(int(mon.replace(',', '').replace('.', '')))))
                                        mon = "".join((mon, ',00'))
                                    if da[0:2].isnumeric() and da[3:5].isnumeric() and da[6:8].isnumeric() and len(da) == 8:
                                        if int(da[3:5]) > 12 or int(da[0:2]) > 31:
                                            messagebox.showinfo(
                                                title='Erreur !!', message="* Le nombre de jour ne doit pas être superieur à 31. \n\n* Le nombre de mois ne doit pas être superieur à 12.", parent=topEdit)
                                            ope = 0
                                            return
                                    else:
                                        messagebox.showinfo(
                                            title='Erreur !!', message="La date doit être sous la forme suivante :\n\n            'JJxMMxAA'", parent=topEdit)
                                        ope = 0
                                        return
                                 
                                    mon=mon.replace('.','')
                                    if len(mon)>4 :
                                        bima=mon[len(mon)-3: len(mon)]
                                        mon=mon[0:len(mon)-3]
                                        for i in range(len(mon)-1,-1,-1) : 
                                            if kok==3 or kok==6 or kok==9 or kok==12 or kok==15 or kok==18 or kok==21:
                                                x=x+'.'
                                                x= x+mon[i]
                                            else : 
                                                x= x+mon[i]
                                            kok=kok+1
                                    if kok!=0 : 
                                        mon=''
                                        for i in range(len(x)-1,-1,-1):
                                            mon=mon+x[i]
                                        mon=mon+bima

                                    hoho = list(da)
                                    hoho[2] = '/'
                                    hoho[5] = '/'
                                    da = ''.join(hoho)
                                    if da != values[0]:
                                        if da[3:5] == '01':
                                            mo = 'JANVIER'
                                        if da[3:5] == '02':
                                            mo = 'FEVRIER'
                                        if da[3:5] == '03':
                                            mo = 'MARS'
                                        if da[3:5] == '04':
                                            mo = 'AVRIL'
                                        if da[3:5] == '05':
                                            mo = 'MAI'
                                        if da[3:5] == '06':
                                            mo = 'JUIN'
                                        if da[3:5] == '07':
                                            mo = 'JUILLET'
                                        if da[3:5] == '08':
                                            mo = 'AOUT'
                                        if da[3:5] == '09':
                                            mo = 'SEPTEMBRE'
                                        if da[3:5] == '10':
                                            mo = 'OCTOBRE'
                                        if da[3:5] == '11':
                                            mo = 'NOVEMBRE'
                                        if da[3:5] == '12':
                                            mo = 'DECEMBRE'
                                    hadighirxhaha = tree.item(
                                        curItem, 'values')[7]
                                    tree.item(curItem, values=(
                                        da, mo, ty, no, mon, dé, mol, att.replace(att[2:4], da[6:8]).replace(att[5:7], da[3:5])))
                                    e9.delete(0, END)
                                    e10.config(state='normal')
                                    e10.delete(0, END)
                                    e12.delete(0, END)
                                    e13.delete(0, END)
                                    e14.delete(0, END)
                                    e15.config(state='normal')
                                    e16.config(state='normal')
                                    e15.delete(0, END)
                                    e16.delete(0, END)
                                    topEdit.destroy()
                                    if values[0].strip() != str(da) or values[2].strip() != str(ty) or values[3].strip() != str(no) or values[4].strip() != str(mon) or values[5].strip() != str(dé):
                                        tree.item(curItem, tags='green')
                                        messagebox.showinfo(
                                            title='Enregistrer', message='Donnée(s) modifiée(s) !')
                                        j = 0
                                        for fi in Dattta:
                                            if hadighirxhaha == fi[7]:
                                                Dattta[j] = [tree.item(curItem, 'values')[0], tree.item(curItem, 'values')[1], tree.item(curItem, 'values')[2], tree.item(curItem, 'values')[
                                                    3], tree.item(curItem, 'values')[4], tree.item(curItem, 'values')[5], tree.item(curItem, 'values')[6], tree.item(curItem, 'values')[7]]
                                            j += 1

                                    ope = 0
                                    return
                                e9.insert(0, values[0])
                                e10.insert(0, values[1])
                                e12.insert(0, values[3])
                                e13.insert(0, values[4])
                                e14.insert(0, values[5])
                                e15.insert(0, values[6])
                                e16.insert(0, values[7])

                                def hantaTchouf():
                                    global ope
                                    ope = 0
                                    e9.delete(0, END)
                                    e10.config(state='normal')
                                    e10.delete(0, END)
                                    e12.delete(0, END)
                                    e13.delete(0, END)
                                    e14.delete(0, END)
                                    e15.config(state='normal')
                                    e16.config(state='normal')
                                    e15.delete(0, END)
                                    e16.delete(0, END)
                                    topEdit.destroy()
                                e16.config(state='disabled')
                                e15.config(state='disabled')
                                e10.config(state='disabled')
                                submitbutton = Button(
                                    topEdit, text='Enregistrer', command=lambda: insertData(my_tree))
                                submitbutton.configure(
                                    font=('Times', 11, 'bold'), bg='green', fg='white')
                                submitbutton.place(x=300, y=350)
                                cancelbutton = Button(
                                    topEdit, text='Annulé', command=lambda: hantaTchouf())
                                cancelbutton.configure(
                                    font=('Times', 11, 'bold'), bg='red', fg='white')
                                cancelbutton.place(x=200, y=350)
                                topEdit.protocol(
                                    "WM_DELETE_WINDOW", hantaTchouf)
                                topEdit.bind(
                                    "<Return>", lambda e: insertData(my_tree))
                                topEdit.bind(
                                    "<Escape>", lambda e: hantaTchouf())
                            else:
                                messagebox.showinfo(
                                    title='!!', message='Merci de selectionner une ligne')
                                ope = 0
                        else:
                            messagebox.showinfo(
                                title='!!', message='Merci de selectionner une seule ligne')
                            ope = 0

                def addData(tree):
                    global Dattta
                    global reloulo
                    global ope
                    if ope == 0:
                        ope = 1

                        def idcode(e):
                            global Dattta
                            l27.configure(
                                text='D-'+e20.get()[6: len(e20.get())]+'-'+e20.get()[3:5]+'-')
                            if e20.get() == '':
                                l27.configure(text='D-AA-MM-')

                        topAdd = Toplevel()
                        topAdd.title("Ajouter")
                        topAdd.geometry("500x400")
                        topAdd.resizable(width=0, height=0)
                        icon = PhotoImage(file='logo-light.png')
                        window.tk.call('wm', 'iconphoto', topAdd._w, icon)
                        l20 = Label(topAdd, text="Date", width=20,
                                    font=('Times', 11, 'bold'))
                        e20 = Entry(
                            topAdd, textvariable=Date, width=25)
                        l20.place(x=50, y=30)
                        e20.place(x=200, y=30)
                        l19 = Label(topAdd, text="Mois", width=20,
                                    font=('Times', 11, 'bold'))
                        e19 = Entry(
                            topAdd, textvariable=Mois, width=25)
                        l19.place(x=50, y=70)
                        e19.place(x=200, y=70)
                        l21 = Label(topAdd, text="Type", width=20,
                                    font=('Times', 11, 'bold'))
                        
                        options = [
                            "VIREMENT PERMANENT",
                            "VIREMENT",
                            "chèque".upper(),
                            "CARTE BANCAIRE",
                            "espèces".upper(),
                        ]
                        
                        # datatype of menu text
                        clicked = StringVar()
                        
                        # initial menu text
                        clicked.set( "VIREMENT" )
                        
                        # Create Dropdown menu
                        drop = OptionMenu( topAdd , clicked , *options )
                        drop.place(x=200, y=105)
                        l21.place(x=50, y=110)
                        l22 = Label(topAdd, text="Nom donneur d'ordre",
                                    width=20, font=('Times', 11, 'bold'))
                        e22 = Entry(
                            topAdd, textvariable=Nom, width=25)
                        l22.place(x=50, y=150)
                        e22.place(x=200, y=150)
                        l23 = Label(topAdd, text="Montant",
                                    width=20, font=('Times', 11, 'bold'))
                        e23 = Entry(
                            topAdd, textvariable=Montant, width=25)
                        l23.place(x=50, y=190)
                        e23.place(x=200, y=190)
                        l24 = Label(topAdd, text="Détail",
                                    width=20, font=('Times', 11, 'bold'))
                        e24 = Entry(
                            topAdd, textvariable=détail, width=25)
                        l24.place(x=50, y=230)
                        e24.place(x=200, y=230)
                        l25 = Label(topAdd, text="Montant en lettre",
                                    width=20, font=('Times', 11, 'bold'))
                        e25 = Entry(
                            topAdd, textvariable=Montant_en_lettre, width=25)
                        l25.place(x=50, y=270)
                        e25.place(x=200, y=270)
                        l26 = Label(topAdd, text="N° Attestation",
                                    width=20, font=('Times', 11, 'bold'))
                        l27 = Label(topAdd, text="D-AA-MM-",
                                    width=20, font=('Times', 11, 'bold'))
                        e26 = Entry(
                            topAdd, textvariable=NAttestation, width=10)
                        l26.place(x=50, y=310)
                        l27.place(x=170, y=315)
                        e26.place(x=300, y=310)
                        e20.focus()
                        e20.bind("<KeyRelease>", idcode)
                        e19.config(state='disabled')
                        e25.config(state='disabled')

                        def adddData(tree):
                            global ope
                            nonlocal e20, e19, e22, e23, e24, e25, e26
                            global Dattta
                            global kn9lb3la
                            global bach
                            global changerowcolo
                            kok=0
                            da = Date.get().strip()
                            mo = Mois.get().strip()
                            ty = clicked.get().upper()
                            no = Nom.get().strip().upper()
                            mon = Montant.get().strip()
                            dé = détail.get().strip()
                            mol = Montant_en_lettre.get().strip()
                            att = "".join(
                                "D-"+da[6:8]+'-'+da[3:5]+'-'+NAttestation.get().strip())
                            if mon == '' or da == '':
                                messagebox.showinfo(
                                    title='Erreur !!', message="la date et le montant sont obligatoire pour enregistrer un paiement.", parent=topAdd)
                                ope = 0
                                return
                            if mon.count(',') == 1:
                                if len(mon)-mon.index(',') != 3:
                                    messagebox.showinfo(
                                        title='Erreur !!', message="Montant invalide, merci de saisir deux chiffres après la virgule.", parent=topAdd)
                                    ope = 0
                                    return
                                else:
                                    for i in mon:
                                        if i == ',' or i == '.':
                                            pass
                                        elif not i.isnumeric():
                                            messagebox.showinfo(
                                                title='Erreur !!', message="Montant invalide", parent=topAdd)
                                            ope = 0
                                            return
                                    po = mon[len(mon)-2] + \
                                        mon[len(mon)-1]
                                    if po == '00':
                                        mol = "".join(chiftolett(
                                            int(int(mon.replace(',', '').replace('.', ''))/100)))
                                    else:
                                        mol = "".join(chiftolett(
                                            int(int(mon.replace(',', '').replace('.', ''))/100)))
                                        mol += 'VIRGULE '
                                        if po == "01" or po == "02" or po == "03" or po == "04" or po == "05" or po == "06" or po == "07" or po == "08" or po == "09":
                                            mol += "ZERO "
                                        mol += "".join(chiftolett(int(po)))
                            else:
                                for i in mon:
                                    if i == '.':
                                        pass
                                    elif not i.isnumeric():
                                        messagebox.showinfo(
                                            title='Erreur !!', message="Montant invalide", parent=topAdd)
                                        ope = 0
                                        return
                                mol = "".join(chiftolett(
                                    int(int(mon.replace(',', '').replace('.', '')))))
                                mon = "".join((mon, ',00'))
                            if da[0:2].isnumeric() and da[3:5].isnumeric() and da[6:8].isnumeric() and len(da) == 8:
                                if int(da[3:5]) > 12 or int(da[0:2]) > 31:
                                    messagebox.showinfo(
                                        title='Erreur !!', message="* Le nombre de jour ne doit pas être superieur à 31. \n\n* Le nombre de mois ne doit pas être superieur à 12.", parent=topAdd)
                                    ope = 0
                                    return
                            else:
                                messagebox.showinfo(
                                    title='Erreur !!', message="La date doit être sous la forme suivante :\n\n            'JJxMMxAA'", parent=topAdd)
                                ope = 0
                                return
                            x=''
                            mon=mon.replace('.','')
                            if len(mon)>4 :
                                bima=mon[len(mon)-3: len(mon)]
                                mon=mon[0:len(mon)-3]
                                for i in range(len(mon)-1,-1,-1) : 
                                    if kok==3 or kok==6 or kok==9 or kok==12 or kok==15 or kok==18 or kok==21:
                                        x=x+'.'
                                        x= x+mon[i]
                                    else : 
                                        x= x+mon[i]
                                    kok=kok+1
                            if kok!=0 : 
                                mon=''
                                for i in range(len(x)-1,-1,-1):
                                    mon=mon+x[i]
                                mon=mon+bima
                            hoho = list(da)
                            hoho[2] = '/'
                            hoho[5] = '/'
                            da = ''.join(hoho)
                            if da[3:5] == '01':
                                mo = 'JANVIER'
                            if da[3:5] == '02':
                                mo = 'FEVRIER'
                            if da[3:5] == '03':
                                mo = 'MARS'
                            if da[3:5] == '04':
                                mo = 'AVRIL'
                            if da[3:5] == '05':
                                mo = 'MAI'
                            if da[3:5] == '06':
                                mo = 'JUIN'
                            if da[3:5] == '07':
                                mo = 'JUILLET'
                            if da[3:5] == '08':
                                mo = 'AOUT'
                            if da[3:5] == '09':
                                mo = 'SEPTEMBRE'
                            if da[3:5] == '10':
                                mo = 'OCTOBRE'
                            if da[3:5] == '11':
                                mo = 'NOVEMBRE'
                            if da[3:5] == '12':
                                mo = 'DECEMBRE'
                            hahia = Dattta[len(
                                Dattta)-1][7][8:len(Dattta[len(Dattta)-1][7])]
                            hahoua = '%0d' % (int(hahia)+1)
                            j = 0
                            try:
                                for li in range(len(Dattta)+1):
                                    if Dattta[li][7][8: len(Dattta[li][7])] == NAttestation.get().strip():
                                        j += 1
                                    if j == 1:
                                        try:
                                            Dattta.insert(
                                                li, [da, mo, ty, no, mon, dé, mol, att])
                                        except:
                                            donothing = 0
                                    if j != 0 and j != 1:
                                        try:
                                            Dattta[li][7] = 'D-'+Dattta[li][0][3:5]+'-' + \
                                                Dattta[li][0][6:len(
                                                    Dattta[li][0])]+'-'+str(li+1)
                                        except:
                                            donothing = 0
                            except:
                                donothing = 0
                            if j != 0:
                                for parent in tree.get_children():
                                    tree.delete(parent)
                                if reloulo == 0:
                                    for i in Dattta:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                    for u in tree.get_children():
                                        if tree.item(u)["values"][7] == att:
                                            tree.focus(u)
                                            tree.selection_set(u)
                                else:
                                    for parent in tree.get_children():
                                        tree.delete(parent)
                                    try:
                                        if bach == 'date':
                                            for i in Dattta:

                                                if kn9lb3la in i[0]:
                                                    
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'mois':
                                            for i in Dattta:
                                                if kn9lb3la in i[1]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'type':
                                            for i in Dattta:
                                                if kn9lb3la in i[2]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'nom':
                                            for i in Dattta:
                                                if kn9lb3la in i[3]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'montant':
                                            for i in Dattta:
                                                if kn9lb3la in i[4]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'détail':
                                            for i in Dattta:
                                                if kn9lb3la in i[5]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'mol':
                                            for i in Dattta:
                                                if kn9lb3la in i[6]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'att':
                                            for i in Dattta:
                                                if kn9lb3la in i[7]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                    except:
                                        donothing = 0                                        
                                if showpdf == 1:
                                    showPDF()
                                e20.delete(0, END)
                                e22.delete(0, END)
                                e23.delete(0, END)
                                e24.delete(0, END)
                                e25.delete(0, END)
                                e26.delete(0, END)
                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                                topAdd.destroy()
                                ope = 0
                                if my_tree.get_children():
                                    for o in range(len(my_tree.get_children())):
                                        if my_tree.item(my_tree.get_children()[o])["values"][7] == att:
                                            my_tree.focus_set()
                                            my_tree.focus(my_tree.get_children()[o])
                                            my_tree.selection_add(my_tree.get_children())
                                            my_tree.selection_set(my_tree.get_children()[o])
                                messagebox.showinfo(
                                    title='Enregistrer', message='Donnée(s) ajoutée(s) !')
                                return    
                            if j == 0:
                                hana = [da, mo, ty, no, mon, dé, mol, "".join(
                                    "D-"+da[6:8]+'-'+da[3:5]+'-'+hahoua)]
                                Dattta.append(hana)
                                for parent in tree.get_children():
                                    tree.delete(parent)
                                if reloulo != 1:
                                    for i in Dattta:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                else:
                                    try:
                                        if bach == 'date':
                                            for i in Dattta:
                                                if kn9lb3la in i[0]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'mois':
                                            for i in Dattta:
                                                if kn9lb3la in i[1]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'type':
                                            for i in Dattta:
                                                if kn9lb3la in i[2]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'nom':
                                            for i in Dattta:
                                                if kn9lb3la in i[3]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'montant':
                                            for i in Dattta:
                                                if kn9lb3la in i[4]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'détail':
                                            for i in Dattta:
                                                if kn9lb3la in i[5]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'mol':
                                            for i in Dattta:
                                                if kn9lb3la in i[6]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                        if bach == 'att':
                                            for i in Dattta:
                                                if kn9lb3la in i[7]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                    except:
                                        donothing = 0
                            if showpdf == 1:
                                showPDF()
                            e20.delete(0, END)
                            e22.delete(0, END)
                            e23.delete(0, END)
                            e24.delete(0, END)
                            e25.delete(0, END)
                            e26.delete(0, END)
                            if changerowcolo == 1:
                                changerowcolo = 0
                                changerowcolor()
                            else:
                                changerowcolo = 1
                                changerowcolor()
                            topAdd.destroy()
                            ope = 0
                            if my_tree.get_children():
                                for o in range(len(my_tree.get_children())):
                                    if my_tree.item(my_tree.get_children()[o])["values"][7] == "".join("D-"+da[6:8]+'-'+da[3:5]+'-'+hahoua):
                                        my_tree.focus_set()
                                        my_tree.focus(my_tree.get_children()[o])
                                        my_tree.selection_add(my_tree.get_children())
                                        my_tree.selection_set(my_tree.get_children()[o])
                            messagebox.showinfo(
                                title='Enregistrer', message='Donnée(s) ajoutée(s) !')
                            return

                        def hantaTchouf():
                            global ope
                            ope = 0
                            e20.delete(0, END)
                            e22.delete(0, END)
                            e23.delete(0, END)
                            e24.delete(0, END)
                            e25.delete(0, END)
                            e26.delete(0, END)
                            topAdd.destroy()
                        submitbutton = Button(
                            topAdd, text='Enregistrer', command=lambda: adddData(my_tree))
                        submitbutton.configure(
                            font=('Times', 11, 'bold'), bg='green', fg='white')
                        submitbutton.place(x=300, y=350)
                        cancelbutton = Button(
                            topAdd, text='Annulé', command=lambda: hantaTchouf())
                        cancelbutton.configure(
                            font=('Times', 11, 'bold'), bg='red', fg='white')
                        cancelbutton.place(x=200, y=350)
                        topAdd.protocol(
                            "WM_DELETE_WINDOW", hantaTchouf)
                        topAdd.bind(
                            "<Return>", lambda e: adddData(my_tree))
                        topAdd.bind(
                            "<Escape>", lambda e: hantaTchouf())
                    else:
                        return

                    return

                def deleteData(tree):
                    global ope
                    global Dattta
                    global kn9lb3la
                    global reloulo
                    global changerowcolo
                    global bach

                    if ope == 0:
                        if len(tree.selection()) > 0:
                            ope = 1
                            sur = messagebox.askquestion(
                                'les données seront perdues !', "êtes-vous sûr de vouloir supprimer cette ligne ?", icon='warning')
                            if sur == 'yes':
                                lololo = tree.selection()
                                for i in lololo:
                                    j = 0
                                    l9it = 0
                                    for y in Dattta:
                                        if l9it == 1:
                                            break
                                        if l9it == 0:
                                            if tree.selection()[tree.selection().index(i)].split()[0] == y[7] or y[7] == tree.selection()[tree.selection().index(i)].split()[len(tree.selection()[tree.selection().index(i)].split())-1]:
                                                tree.delete(tree.selection()[
                                                            tree.selection().index(i)])
                                                Dattta.remove(Dattta[j])
                                                l9it = 1
                                        j += 1
                                for i in lololo:
                                    for lpo in range(len(Dattta)):
                                        try:
                                            Dattta[lpo][7] = 'D-'+Dattta[lpo][0][6:len(
                                                Dattta[lpo][0])]+'-' + Dattta[lpo][0][3:5]+'-'+str(lpo+1)
                                        except:
                                            donothing = 0

                                for parent in tree.get_children():
                                    tree.delete(parent)
                                if reloulo == 0:
                                    for i in Dattta:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                            mytag = 'red'
                                        tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                    if changerowcolo == 1:
                                        changerowcolo = 0
                                        changerowcolor()
                                    else:
                                        changerowcolo = 1
                                        changerowcolor()
                                else:

                                    try:
                                        if bach == 'date':
                                            for i in Dattta:
                                                if kn9lb3la in i[0]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                            if changerowcolo == 1:
                                                changerowcolo = 0
                                                changerowcolor()
                                            else:
                                                changerowcolo = 1
                                                changerowcolor()
                                        if bach == 'mois':
                                            for i in Dattta:
                                                if kn9lb3la in i[1]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                            if changerowcolo == 1:
                                                changerowcolo = 0
                                                changerowcolor()
                                            else:
                                                changerowcolo = 1
                                                changerowcolor()
                                        if bach == 'type':
                                            for i in Dattta:
                                                if kn9lb3la in i[2]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                            if changerowcolo == 1:
                                                changerowcolo = 0
                                                changerowcolor()
                                            else:
                                                changerowcolo = 1
                                                changerowcolor()
                                        if bach == 'nom':
                                            for i in Dattta:
                                                if kn9lb3la in i[3]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                            if changerowcolo == 1:
                                                changerowcolo = 0
                                                changerowcolor()
                                            else:
                                                changerowcolo = 1
                                                changerowcolor()
                                        if bach == 'montant':
                                            for i in Dattta:
                                                if kn9lb3la in i[4]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                            if changerowcolo == 1:
                                                changerowcolo = 0
                                                changerowcolor()
                                            else:
                                                changerowcolo = 1
                                                changerowcolor()
                                        if bach == 'détail':
                                            for i in Dattta:
                                                if kn9lb3la in i[5]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                            if changerowcolo == 1:
                                                changerowcolo = 0
                                                changerowcolor()
                                            else:
                                                changerowcolo = 1
                                                changerowcolor()
                                        if bach == 'mol':
                                            for i in Dattta:
                                                if kn9lb3la in i[6]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                            if changerowcolo == 1:
                                                changerowcolo = 0
                                                changerowcolor()
                                            else:
                                                changerowcolo = 1
                                                changerowcolor()
                                        if bach == 'att':
                                            for i in Dattta:
                                                if kn9lb3la in i[7]:
                                                    mytag = 'normal'
                                                    if i[2] == 'CARTE BANCAIRE':
                                                        mytag = 'blue'
                                                    if i[2] != 'CARTE BANCAIRE' and (i[3] == '' or i[3] == ' '):
                                                        mytag = 'red'
                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                    mytag = 'normal'
                                            if changerowcolo == 1:
                                                changerowcolo = 0
                                                changerowcolor()
                                            else:
                                                changerowcolo = 1
                                                changerowcolor()
                                    except:
                                        donothing = 0

                            ope = 0
                            children = my_tree.get_children()
                            if children:
                                my_tree.focus(children[0])
                                my_tree.selection_set(children[0])
                                my_tree.selection_add(children[0])
                        else:
                            messagebox.showinfo(
                                title='!!', message='Merci de selectionner les lignes à supprimer')

                def menuRecherch():
                    global rech
                    if rech == 0:
                        rech = 1
                        global deb

                        def check1():
                            global deb
                            deb = 1
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            c1.select()
                            checkVar2 = BooleanVar(value=False)
                            checkVar3 = BooleanVar(value=False)
                            checkVar4 = BooleanVar(value=False)
                            checkVar5 = BooleanVar(value=False)
                            checkVar6 = BooleanVar(value=False)
                            checkVar7 = BooleanVar(value=False)
                            checkVar8 = BooleanVar(value=False)
                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.config(state='normal')
                            e2.config(state='disabled')
                            e3.config(state='disabled')
                            e4.config(state='disabled')
                            e5.config(state='disabled')
                            e6.config(state='disabled')
                            e7.config(state='disabled')
                            e8.config(state='disabled')
                            topSearch.destroy()
                            global rech
                            rech = 0
                            menuRecherch()
                            return

                        def check2():
                            global deb
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            deb = 2
                            checkVar1 = BooleanVar(value=False)
                            checkVar2 = BooleanVar(value=True)
                            checkVar3 = BooleanVar(value=False)
                            checkVar4 = BooleanVar(value=False)
                            checkVar5 = BooleanVar(value=False)
                            checkVar6 = BooleanVar(value=False)
                            checkVar7 = BooleanVar(value=False)
                            checkVar8 = BooleanVar(value=False)
                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.config(state='disabled')
                            e2.config(state='normal')
                            e3.config(state='disabled')
                            e4.config(state='disabled')
                            e5.config(state='disabled')
                            e6.config(state='disabled')
                            e7.config(state='disabled')
                            e8.config(state='disabled')
                            topSearch.destroy()
                            global rech
                            rech = 0
                            menuRecherch()
                            return

                        def check3():
                            global deb
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            checkVar1 = BooleanVar(value=False)
                            checkVar2 = BooleanVar(value=False)
                            checkVar3 = BooleanVar(value=True)
                            checkVar4 = BooleanVar(value=False)
                            deb = 3
                            checkVar5 = BooleanVar(value=False)
                            checkVar6 = BooleanVar(value=False)
                            checkVar7 = BooleanVar(value=False)
                            checkVar8 = BooleanVar(value=False)

                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.config(state='disabled')
                            e2.config(state='disabled')
                            e3.config(state='normal')
                            e4.config(state='disabled')
                            e5.config(state='disabled')
                            e6.config(state='disabled')
                            e7.config(state='disabled')
                            e8.config(state='disabled')
                            topSearch.destroy()
                            global rech
                            rech = 0
                            menuRecherch()
                            return

                        def check4():
                            global deb
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            checkVar1 = BooleanVar(value=False)
                            deb = 4
                            checkVar2 = BooleanVar(value=False)
                            checkVar3 = BooleanVar(value=False)
                            checkVar4 = BooleanVar(value=True)
                            checkVar5 = BooleanVar(value=False)
                            checkVar6 = BooleanVar(value=False)
                            checkVar7 = BooleanVar(value=False)
                            checkVar8 = BooleanVar(value=False)
                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.config(state='disabled')
                            e2.config(state='disabled')
                            e3.config(state='disabled')
                            e4.config(state='normal')
                            e5.config(state='disabled')
                            e6.config(state='disabled')
                            e7.config(state='disabled')
                            e8.config(state='disabled')
                            topSearch.destroy()
                            global rech
                            rech = 0
                            menuRecherch()
                            return

                        def check5():
                            global deb
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            checkVar1 = BooleanVar(value=False)
                            checkVar2 = BooleanVar(value=False)
                            deb = 5
                            checkVar3 = BooleanVar(value=False)
                            checkVar4 = BooleanVar(value=False)
                            checkVar5 = BooleanVar(value=True)
                            checkVar6 = BooleanVar(value=False)
                            checkVar7 = BooleanVar(value=False)
                            checkVar8 = BooleanVar(value=False)
                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.config(state='disabled')
                            e2.config(state='disabled')
                            e3.config(state='disabled')
                            e4.config(state='disabled')
                            e5.config(state='normal')
                            e6.config(state='disabled')
                            e7.config(state='disabled')
                            e8.config(state='disabled')
                            topSearch.destroy()
                            global rech
                            rech = 0
                            menuRecherch()
                            return

                        def check6():
                            global deb
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            checkVar1 = BooleanVar(value=False)
                            checkVar2 = BooleanVar(value=False)
                            checkVar3 = BooleanVar(value=False)
                            checkVar4 = BooleanVar(value=False)
                            checkVar5 = BooleanVar(value=False)
                            deb = 6
                            checkVar6 = BooleanVar(value=True)
                            checkVar7 = BooleanVar(value=False)
                            checkVar8 = BooleanVar(value=False)
                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.config(state='disabled')
                            e2.config(state='disabled')
                            e3.config(state='disabled')
                            e4.config(state='disabled')
                            e5.config(state='disabled')
                            e6.config(state='normal')
                            e7.config(state='disabled')
                            e8.config(state='disabled')
                            topSearch.destroy()
                            global rech
                            rech = 0
                            menuRecherch()
                            return

                        def check7():
                            global deb
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            checkVar1 = BooleanVar(value=False)
                            checkVar2 = BooleanVar(value=False)
                            checkVar3 = BooleanVar(value=False)
                            checkVar4 = BooleanVar(value=False)
                            checkVar5 = BooleanVar(value=False)
                            checkVar6 = BooleanVar(value=False)
                            deb = 7
                            checkVar7 = BooleanVar(value=True)
                            checkVar8 = BooleanVar(value=False)
                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.config(state='disabled')
                            e2.config(state='disabled')
                            e3.config(state='disabled')
                            e4.config(state='disabled')
                            e5.config(state='disabled')
                            e6.config(state='disabled')
                            e7.config(state='normal')
                            e8.config(state='disabled')
                            topSearch.destroy()
                            global rech
                            rech = 0
                            menuRecherch()
                            return

                        def check8():
                            global deb
                            global rech
                            deb = 0
                            rech = 0
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            checkVar1 = BooleanVar(value=False)
                            checkVar2 = BooleanVar(value=False)
                            checkVar3 = BooleanVar(value=False)
                            checkVar4 = BooleanVar(value=False)
                            checkVar5 = BooleanVar(value=False)
                            checkVar6 = BooleanVar(value=False)
                            checkVar7 = BooleanVar(value=False)
                            checkVar8 = BooleanVar(value=True)
                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.config(state='disabled')
                            e2.config(state='disabled')
                            e3.config(state='disabled')
                            e4.config(state='disabled')
                            e5.config(state='disabled')
                            e6.config(state='disabled')
                            e7.config(state='disabled')
                            e8.config(state='normal')
                            topSearch.destroy()
                            menuRecherch()
                            return
                        topSearch = Toplevel()
                        topSearch.title("Recherche")
                        topSearch.geometry("1000x400")
                        icon = PhotoImage(file='logo-light.png')
                        window.tk.call('wm', 'iconphoto', topSearch._w, icon)
                        loubil = Listbox(topSearch, bg='white',
                                         activestyle='dotbox', justify="center")
                        louou = Label(topSearch, text='',  font=('Times', 18))
                        loubil.place(width=550, height=300, x=430, y=30)
                        topSearch.resizable(width=0, height=0)
                        loubil.delete(0, END)
                        c1 = Checkbutton(topSearch, text='   Date                                          ', variable=checkVar1,
                                         font=('Times', 11, 'bold'),
                                         onvalue=1, offvalue=0, command=lambda: check1())
                        c1.place(x=20, y=30)
                        e1 = Entry(
                            topSearch, textvariable=Date, width=25)
                        e1.place(x=200, y=30)
                        e1.config(state='disabled')
                        if deb == 1:
                            louou.configure(text='DATES')
                            louou.place(width=550, height=20, x=430, y=10)
                            e1.config(state='normal')
                            e1.focus()
                            c1.select()
                            topSearch.title("Recherche par date")
                            for parent in my_tree.get_children():
                                x = my_tree.item(parent)["values"][0]
                                loubil.insert(END, x)
                        c2 = Checkbutton(topSearch, text='   Mois                                          ', variable=checkVar2,
                                         font=('Times', 11, 'bold'),
                                         onvalue=1, offvalue=0, command=lambda: check2())
                        c2.place(x=20, y=70)
                        e2 = Entry(
                            topSearch, textvariable=Mois, width=25)
                        e2.place(x=200, y=70)
                        e2.config(state='disabled')
                        if deb == 2:
                            louou.configure(text='MOIS')
                            louou.place(width=550, height=20, x=430, y=10)
                            e2.config(state='normal')
                            e2.focus()
                            c2.select()
                            topSearch.title("Recherche par mois")
                            for parent in my_tree.get_children():
                                x = my_tree.item(parent)["values"][1]
                                loubil.insert(END, x)
                        c3 = Checkbutton(topSearch, text='   Type                                        ', variable=checkVar3,
                                         font=('Times', 11, 'bold'),
                                         onvalue=1, offvalue=0, command=lambda: check3())
                        c3.place(x=20, y=110)
                        e3 = Entry(
                            topSearch, textvariable=Type, width=25)
                        e3.place(x=200, y=110)
                        e3.config(state='disabled')
                        if deb == 3:
                            louou.configure(text='TYPES')
                            louou.place(width=550, height=20, x=430, y=10)
                            e3.config(state='normal')
                            e3.focus()
                            c3.select()
                            topSearch.title("Recherche par type")
                            for parent in my_tree.get_children():
                                x = my_tree.item(parent)["values"][2]
                                loubil.insert(END, x)
                        c4 = Checkbutton(topSearch, text="   Nom donneur d'ordre                  ", variable=checkVar4,
                                         font=('Times', 11, 'bold'),
                                         onvalue=1, offvalue=0, command=lambda: check4())
                        c4.place(x=20, y=150)
                        e4 = Entry(topSearch, textvariable=Nom, width=25)
                        e4.place(x=200, y=150)
                        e4.config(state='disabled')
                        if deb == 4:
                            louou.configure(text='NOMS')
                            louou.place(width=550, height=20, x=430, y=10)
                            e4.config(state='normal')
                            e4.focus()
                            c4.select()
                            topSearch.title("Recherche par nom")
                            for parent in my_tree.get_children():
                                x = my_tree.item(parent)["values"][3]
                                loubil.insert(END, x)
                        c5 = Checkbutton(topSearch, text="   Montant                                  ", variable=checkVar5,
                                         font=('Times', 11, 'bold'),
                                         onvalue=1, offvalue=0, command=lambda: check5())
                        c5.place(x=20, y=190)
                        e5 = Entry(
                            topSearch, textvariable=Montant, width=25)
                        e5.place(x=200, y=190)
                        e5.config(state='disabled')
                        if deb == 5:
                            louou.configure(text='MONTANTS')
                            louou.place(width=550, height=20, x=430, y=10)
                            e5.config(state='normal')
                            e5.focus()
                            c5.select()
                            topSearch.title("Recherche par montant")
                            for parent in my_tree.get_children():
                                x = my_tree.item(parent)["values"][4]
                                loubil.insert(END, x)
                        c6 = Checkbutton(topSearch, text="   Détail                                   ", variable=checkVar6,
                                         font=('Times', 11, 'bold'),
                                         onvalue=1, offvalue=0, command=lambda: check6())
                        c6.place(x=20, y=230)
                        e6 = Entry(
                            topSearch, textvariable=détail, width=25)
                        e6.place(x=200, y=230)
                        e6.config(state='disabled')
                        if deb == 6:
                            louou.configure(text='DETAIL')
                            louou.place(width=550, height=20, x=430, y=10)
                            e6.config(state='normal')
                            e6.focus()
                            c6.select()
                            topSearch.title("Recherche par détail")
                            for parent in my_tree.get_children():
                                x = my_tree.item(parent)["values"][5]
                                loubil.insert(END, x)
                        c7 = Checkbutton(topSearch, text="   Montant en lettre              ", variable=checkVar7,
                                         font=('Times', 11, 'bold'),
                                         onvalue=1, offvalue=0, command=lambda: check7())
                        c7.place(x=20, y=270)
                        e7 = Entry(
                            topSearch, textvariable=Montant_en_lettre, width=25)
                        e7.place(x=200, y=270)
                        e7.config(state='disabled')
                        if deb == 7:
                            louou.configure(text='MONTANTS EN LETTRE')
                            louou.place(width=550, height=20, x=430, y=10)
                            e7.config(state='normal')
                            e7.focus()
                            c7.select()
                            topSearch.title("Recherche par montent en lettre")
                            for parent in my_tree.get_children():
                                x = my_tree.item(parent)["values"][6]
                                loubil.insert(END, x)
                        c8 = Checkbutton(topSearch, text="   N° Attestation                  ", variable=checkVar8,
                                         font=('Times', 11, 'bold'),
                                         onvalue=1, offvalue=0, command=lambda: check8())
                        c8.place(x=20, y=310)
                        e8 = Entry(
                            topSearch, textvariable=NAttestation, width=25)
                        e8.place(x=200, y=310)
                        e8.config(state='disabled')
                        if deb == 0:
                            louou.configure(text="NUMEROS D'ATTESTATIONS")
                            louou.place(width=550, height=20, x=430, y=10)
                            e8.config(state='normal')
                            e8.focus()
                            c8.select()
                            topSearch.title(
                                "Recherche par numero d'attestation")
                            for parent in my_tree.get_children():
                                x = my_tree.item(parent)["values"][7]
                                loubil.insert(END, x)

                        def stopSearch():
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            global deb
                            deb = 0
                            global rech
                            rech = 0
                            checkVar1 = BooleanVar(value=False)
                            checkVar2 = BooleanVar(value=False)
                            checkVar3 = BooleanVar(value=False)
                            checkVar4 = BooleanVar(value=False)
                            checkVar5 = BooleanVar(value=False)
                            checkVar6 = BooleanVar(value=False)
                            checkVar7 = BooleanVar(value=False)
                            checkVar8 = BooleanVar(value=True)
                            topSearch.destroy()

                        def autofill(e):
                            e1.delete(0, END)
                            e2.delete(0, END)
                            e3.delete(0, END)
                            e4.delete(0, END)
                            e5.delete(0, END)
                            e6.delete(0, END)
                            e7.delete(0, END)
                            e8.delete(0, END)
                            e1.insert(0, loubil.get(loubil.curselection()))
                            e2.insert(0, loubil.get(loubil.curselection()))
                            e3.insert(0, loubil.get(loubil.curselection()))
                            e4.insert(0, loubil.get(loubil.curselection()))
                            e5.insert(0, loubil.get(loubil.curselection()))
                            e6.insert(0, loubil.get(loubil.curselection()))
                            e7.insert(0, loubil.get(loubil.curselection()))
                            e8.insert(0, loubil.get(loubil.curselection()))

                        def tl31(e):
                            global Dattta
                            typed = e1.get()
                            if typed == '':
                                for parent in Dattta:
                                    loubil.insert(END, parent[0])
                            else:
                                loubil.delete(0, END)
                                for parent in Dattta:
                                    if typed.upper() in parent[0].upper():
                                        loubil.insert(END, parent[0])

                        def tl32(e):
                            typed = e2.get()
                            if typed == '':
                                for parent in Dattta:
                                    loubil.insert(END, parent[1])
                            else:
                                loubil.delete(0, END)
                                for parent in Dattta:
                                    if typed.upper() in parent[1].upper():
                                        loubil.insert(END, parent[1])

                        def tl33(e):
                            typed = e3.get()
                            if typed == '':
                                for parent in Dattta:
                                    loubil.insert(END, parent[2])
                            else:
                                loubil.delete(0, END)
                                for parent in Dattta:
                                    if typed.upper() in parent[2].upper():
                                        loubil.insert(END, parent[2])

                        def tl34(e):
                            typed = e4.get()
                            if typed == '':
                                for parent in Dattta:
                                    loubil.insert(END, parent[3])
                            else:
                                loubil.delete(0, END)
                                for parent in Dattta:
                                    if typed.upper() in parent[3].upper():
                                        loubil.insert(END, parent[3])

                        def tl35(e):
                            typed = e5.get()
                            if typed == '':
                                for parent in Dattta:
                                    loubil.insert(END, parent[4])
                            else:
                                loubil.delete(0, END)
                                for parent in Dattta:
                                    if typed.upper() in parent[4].upper():
                                        loubil.insert(END, parent[4])

                        def tl36(e):
                            typed = e6.get()
                            if typed == '':
                                for parent in Dattta:
                                    loubil.insert(END, parent[5])
                            else:
                                loubil.delete(0, END)
                                for parent in Dattta:
                                    if typed.upper() in parent[5].upper():
                                        loubil.insert(END, parent[5])

                        def tl37(e):
                            typed = e7.get()
                            if typed == '':
                                for parent in Dattta:
                                    loubil.insert(END, parent[6])
                            else:
                                loubil.delete(0, END)
                                for parent in Dattta:
                                    if typed.upper() in parent[6].upper():
                                        loubil.insert(END, parent[6])

                        def tl38(e):
                            typed = e8.get()
                            if typed == '':
                                for parent in Dattta:
                                    loubil.insert(END, parent[7])
                            else:
                                loubil.delete(0, END)
                                for parent in Dattta:
                                    if typed.upper() in parent[7].upper():
                                        loubil.insert(END, parent[7])

                        loubil.bind("<<ListboxSelect>>", autofill)
                        e1.bind("<KeyRelease>", tl31)
                        e2.bind("<KeyRelease>", tl32)
                        e3.bind("<KeyRelease>", tl33)
                        e4.bind("<KeyRelease>", tl34)
                        e5.bind("<KeyRelease>", tl35)
                        e6.bind("<KeyRelease>", tl36)
                        e7.bind("<KeyRelease>", tl37)
                        e8.bind("<KeyRelease>", tl38)

                        def searchData(tree):
                            global rech
                            global deb
                            global Dattta
                            global ka
                            global reloulo
                            global reloulou
                            global bach
                            global kn9lb3la
                            global rj3Data
                            global changerowcolo
                            global checkVar1, checkVar2, checkVar3, checkVar4, checkVar5, checkVar6, checkVar7, checkVar8
                            movedown.config(state='disabled')
                            moveup.config(state='disabled')

                            def rj3Data(tree):
                                global showpdf
                                global Dattta
                                global ka
                                global reloulo
                                global kn9lb3la
                                global bach
                                global changerowcolo
                                global reloulou
                                movedown.config(state='normal')
                                moveup.config(state='normal')
                                ka = 0
                                reloulo = 0
                                reloulou.place_forget()
                                for parent in tree.get_children():
                                    tree.delete(parent)
                                for i in Dattta:
                                    mytag = 'normal'
                                    if i[2] == 'CARTE BANCAIRE':
                                        mytag = 'blue'
                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                        mytag = 'red'
                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                    mytag = 'normal'

                                if showpdf == 1:
                                    showPDF()
                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()

                            if deb == 1:
                                if showpdf == 1:
                                    showPDF()
                                ha3lachkt9lbe = e1.get().upper()
                                kn9lb3la = ha3lachkt9lbe
                                e1.delete(0, END)
                                deb = 0
                                rech = 0
                                checkVar1 = BooleanVar(value=False)
                                checkVar8 = BooleanVar(value=True)
                                for parent in tree.get_children():
                                    tree.delete(parent)

                                for i in Dattta:
                                    if kn9lb3la in i[0]:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                        ka = 1
                                        bach = 'date'
                                if ka == 1:
                                    stopSearch()

                                    if reloulo == 0:
                                        reloulou = Button(
                                            frame1, text='<-', command=lambda: rj3Data(my_tree))
                                        reloulou.place(x=0, y=137)
                                        reloulo = 1
                                        
                                else:
                                    rech = 1
                                    c1.select()
                                    e1.focus()
                                    messagebox.showinfo(
                                        title=kn9lb3la, message="Pas de données !", parent=topSearch)

                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                            if deb == 2:
                                if showpdf == 1:
                                    showPDF()
                                kn9lb3la = e2.get().upper()
                                e2.delete(0, END)
                                deb = 0
                                rech = 0
                                checkVar1 = BooleanVar(value=False)
                                checkVar8 = BooleanVar(value=True)
                                for parent in tree.get_children():
                                    tree.delete(parent)

                                for i in Dattta:
                                    if kn9lb3la in i[1]:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                        ka = 1
                                        bach = 'mois'
                                if ka == 1:
                                    stopSearch()
                                    if reloulo == 0:
                                        reloulou = Button(
                                            frame1, text='<-', command=lambda: rj3Data(my_tree))
                                        reloulou.place(x=0, y=137)
                                        reloulo = 1
                                else:
                                    rech = 1
                                    c2.select()
                                    e2.focus()
                                    messagebox.showinfo(
                                        title=kn9lb3la, message="Pas de données !", parent=topSearch)

                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                            if deb == 3:
                                if showpdf == 1:
                                    showPDF()
                                kn9lb3la = e3.get().upper()
                                e3.delete(0, END)
                                deb = 0
                                rech = 0
                                checkVar1 = BooleanVar(value=False)
                                checkVar8 = BooleanVar(value=True)
                                for parent in tree.get_children():
                                    tree.delete(parent)

                                for i in Dattta:
                                    if kn9lb3la in i[2]:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                        ka = 1
                                        bach = 'type'
                                if ka == 1:
                                    stopSearch()
                                    if reloulo == 0:
                                        reloulou = Button(
                                            frame1, text='<-', command=lambda: rj3Data(my_tree))
                                        reloulou.place(x=0, y=137)
                                        reloulo = 1
                                else:
                                    rech = 1
                                    c3.select()
                                    e3.focus()
                                    messagebox.showinfo(
                                        title=kn9lb3la, message="Pas de données !", parent=topSearch)

                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                            if deb == 4:
                                if showpdf == 1:
                                    showPDF()
                                kn9lb3la = e4.get().upper()
                                e4.delete(0, END)
                                deb = 0
                                rech = 0
                                checkVar1 = BooleanVar(value=False)
                                checkVar8 = BooleanVar(value=True)
                                for parent in tree.get_children():
                                    tree.delete(parent)

                                for i in Dattta:
                                    if kn9lb3la in i[3]:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                        ka = 1
                                        bach = 'nom'
                                if ka == 1:
                                    stopSearch()
                                    if reloulo == 0:
                                        reloulou = Button(
                                            frame1, text='<-', command=lambda: rj3Data(my_tree))
                                        reloulou.place(x=0, y=137)
                                        reloulo = 1
                                else:
                                    rech = 1
                                    c4.select()
                                    e4.focus()
                                    messagebox.showinfo(
                                        title=kn9lb3la, message="Pas de données !", parent=topSearch)

                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                            if deb == 5:
                                if showpdf == 1:
                                    showPDF()
                                kn9lb3la = e5.get().upper()
                                e5.delete(0, END)
                                deb = 0
                                rech = 0
                                checkVar1 = BooleanVar(value=False)
                                checkVar8 = BooleanVar(value=True)
                                for parent in tree.get_children():
                                    tree.delete(parent)

                                for i in Dattta:
                                    if kn9lb3la in i[4]:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                        ka = 1
                                        bach = 'montant'
                                if ka == 1:
                                    stopSearch()
                                    if reloulo == 0:
                                        reloulou = Button(
                                            frame1, text='<-', command=lambda: rj3Data(my_tree))
                                        reloulou.place(x=0, y=137)
                                        reloulo = 1
                                else:
                                    rech = 1
                                    c5.select()
                                    e5.focus()
                                    messagebox.showinfo(
                                        title=kn9lb3la, message="Pas de données !", parent=topSearch)

                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                            if deb == 6:
                                if showpdf == 1:
                                    showPDF()
                                kn9lb3la = e6.get().upper()
                                e6.delete(0, END)
                                deb = 0
                                rech = 0
                                checkVar1 = BooleanVar(value=False)
                                checkVar8 = BooleanVar(value=True)
                                for parent in tree.get_children():
                                    tree.delete(parent)

                                for i in Dattta:
                                    if kn9lb3la in i[5]:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                        ka = 1
                                        bach = 'détail'
                                if ka == 1:
                                    stopSearch()
                                    if reloulo == 0:
                                        reloulou = Button(
                                            frame1, text='<-', command=lambda: rj3Data(my_tree))
                                        reloulou.place(x=0, y=137)
                                        reloulo = 1
                                else:
                                    rech = 1
                                    c6.select()
                                    e6.focus()
                                    messagebox.showinfo(
                                        title=kn9lb3la, message="Pas de données !", parent=topSearch)

                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                            if deb == 7:
                                if showpdf == 1:
                                    showPDF()
                                kn9lb3la = e7.get().upper()
                                e7.delete(0, END)
                                deb = 0
                                rech = 0
                                checkVar1 = BooleanVar(value=False)
                                checkVar8 = BooleanVar(value=True)
                                for parent in tree.get_children():
                                    tree.delete(parent)

                                for i in Dattta:
                                    if kn9lb3la in i[6]:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                        ka = 1
                                        bach = 'mol'
                                if ka == 1:
                                    stopSearch()
                                    if reloulo == 0:
                                        reloulou = Button(
                                            frame1, text='<-', command=lambda: rj3Data(my_tree))
                                        reloulou.place(x=0, y=137)
                                        reloulo = 1
                                else:
                                    rech = 1
                                    c7.select()
                                    e7.focus()
                                    messagebox.showinfo(
                                        title=kn9lb3la, message="Pas de données !", parent=topSearch)

                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                            if deb == 0:
                                if showpdf == 1:
                                    showPDF()
                                kn9lb3la = e8.get().upper()
                                e8.delete(0, END)
                                deb = 0
                                rech = 0
                                checkVar1 = BooleanVar(value=False)
                                checkVar8 = BooleanVar(value=True)
                                for parent in tree.get_children():
                                    tree.delete(parent)

                                for i in Dattta:
                                    if kn9lb3la in i[7]:
                                        mytag = 'normal'
                                        if i[2] == 'CARTE BANCAIRE':
                                            mytag = 'blue'
                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                            mytag = 'red'
                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                        mytag = 'normal'
                                        ka = 1
                                        bach = 'att'
                                if ka == 1:
                                    stopSearch()
                                    if reloulo == 0:
                                        reloulou = Button(
                                            frame1, text='<-', command=lambda: rj3Data(my_tree))
                                        reloulou.place(x=0, y=137)
                                        reloulo = 1
                                else:
                                    rech = 1
                                    c8.select()
                                    e8.focus()
                                    messagebox.showinfo(
                                        title=kn9lb3la, message="Pas de données !", parent=topSearch)

                                if changerowcolo == 1:
                                    changerowcolo = 0
                                    changerowcolor()
                                else:
                                    changerowcolo = 1
                                    changerowcolor()
                            return
                        ch = Button(
                            topSearch, text='Chercher', command=lambda: searchData(my_tree))
                        ch.configure(
                            font=('Times', 11, 'bold'), bg='green', fg='white')
                        ch.place(x=300, y=350)
                        anul = Button(
                            topSearch, text='Annulé')
                        anul.configure(
                            font=('Times', 11, 'bold'), bg='red', fg='white')
                        anul.place(x=200, y=350)
                        topSearch.protocol("WM_DELETE_WINDOW", stopSearch)
                        topSearch.bind(
                            "<Return>", lambda e: searchData(my_tree))
                        topSearch.bind(
                            "<Escape>", lambda e: stopSearch())
                    return

                def changerowcolor():
                    global mytag
                    global bach
                    global reloulo
                    global changerowcolo
                    if changerowcolo == 0:
                        for parent in my_tree.get_children():
                            my_tree.delete(parent)
                        if reloulo == 0:
                            for i in Dattta:
                                if mytag == 'normal':
                                    mytag = 'gray'
                                else:
                                    mytag = 'normal'
                                if i[2] == 'CARTE BANCAIRE':
                                    mytag = 'blue'
                                if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                    mytag = 'red'
                                my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                    i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                        else:
                            try:
                                if bach == 'date':
                                    for i in Dattta:
                                        if kn9lb3la in i[0]:
                                            if mytag == 'normal':
                                                mytag = 'gray'
                                            else:
                                                mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'mois':
                                    for i in Dattta:
                                        if kn9lb3la in i[1]:
                                            if mytag == 'normal':
                                                mytag = 'gray'
                                            else:
                                                mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'type':
                                    for i in Dattta:
                                        if kn9lb3la in i[2]:
                                            if mytag == 'normal':
                                                mytag = 'gray'
                                            else:
                                                mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'nom':
                                    for i in Dattta:
                                        if kn9lb3la in i[3]:
                                            if mytag == 'normal':
                                                mytag = 'gray'
                                            else:
                                                mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'montant':
                                    for i in Dattta:
                                        if kn9lb3la in i[4]:
                                            if mytag == 'normal':
                                                mytag = 'gray'
                                            else:
                                                mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'détail':
                                    for i in Dattta:
                                        if kn9lb3la in i[5]:
                                            if mytag == 'normal':
                                                mytag = 'gray'
                                            else:
                                                mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'mol':
                                    for i in Dattta:
                                        if kn9lb3la in i[6]:
                                            if mytag == 'normal':
                                                mytag = 'gray'
                                            else:
                                                mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'att':
                                    for i in Dattta:
                                        if kn9lb3la in i[7]:
                                            if mytag == 'normal':
                                                mytag = 'gray'
                                            else:
                                                mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                            except:
                                donothing = 0
                        changerowcolo = 1

                    else:
                        for parent in my_tree.get_children():
                            my_tree.delete(parent)
                        if reloulo == 0:
                            for i in Dattta:
                                mytag = 'normal'
                                if i[2] == 'CARTE BANCAIRE':
                                    mytag = 'blue'
                                if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                    mytag = 'red'
                                my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                    i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                        else:
                            try:
                                if bach == 'date':
                                    for i in Dattta:
                                        if kn9lb3la in i[0]:
                                            mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'mois':
                                    for i in Dattta:
                                        if kn9lb3la in i[1]:
                                            mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'type':
                                    for i in Dattta:
                                        if kn9lb3la in i[2]:
                                            mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'nom':
                                    for i in Dattta:
                                        if kn9lb3la in i[3]:
                                            mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'montant':
                                    for i in Dattta:
                                        if kn9lb3la in i[4]:
                                            mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'détail':
                                    for i in Dattta:
                                        if kn9lb3la in i[5]:
                                            mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'mol':
                                    for i in Dattta:
                                        if kn9lb3la in i[6]:
                                            mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                if bach == 'att':
                                    for i in Dattta:
                                        if kn9lb3la in i[7]:
                                            mytag = 'normal'
                                            if i[2] == 'CARTE BANCAIRE':
                                                mytag = 'blue'
                                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                mytag = 'red'
                                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                            except:
                                donothing = 0
                        changerowcolo = 0
                    return

                def selectibabahom():
                    global mselec
                    if mselec == 0:
                        if children:
                            my_tree.selection_set(my_tree.get_children()[0])
                            my_tree.focus_set()
                            my_tree.focus(my_tree.get_children()[0])
                            my_tree.selection_add(my_tree.get_children())
                            mselec = 1
                    else:
                        for item in my_tree.selection():
                            my_tree.selection_remove(item)
                        mselec = 0

                def up():
                    global ope
                    ope = 1
                    Dattta = []
                    x = []
                    rows = my_tree.selection()
                    if len(my_tree.selection()) < 1:
                        messagebox.showinfo(
                            title='Erreur !!', message='Merci de selectionner une ou plusieurs lignes. ')
                    else:
                        for row in rows:
                            my_tree.move(row, my_tree.parent(
                                row), my_tree.index(row)-1)
                            x.append(my_tree.index(row)-1)
                        for parent in my_tree.get_children():
                            Dattta.append(my_tree.item(parent)["values"])
                        for lpo in range(len(Dattta)):
                            try:
                                Dattta[lpo][7] = 'D-'+Dattta[lpo][0][6:len(
                                    Dattta[lpo][0])]+'-' + Dattta[lpo][0][3:5]+'-'+str(lpo+1)
                            except:
                                donothing = 0
                        for parent in my_tree.get_children():
                            my_tree.delete(parent)
                        for i in Dattta:
                            mytag = 'normal'
                            if i[2] == 'CARTE BANCAIRE':
                                mytag = 'blue'
                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                mytag = 'red'
                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                            mytag = 'normal'
                        for i in x:
                            if children:
                                my_tree.selection_add(
                                    my_tree.get_children()[i+1])
                                my_tree.focus_set()
                                my_tree.focus(my_tree.get_children()[i+1])
                    ope = 0
                    return

                def down():
                    global ope
                    ope = 1
                    Dattta = []
                    x = []
                    rows = my_tree.selection()
                    if len(my_tree.selection()) < 1:
                        messagebox.showinfo(
                            title='Erreur !!', message='Merci de selectionner une ou plusieurs lignes. ')
                    else:
                        for row in reversed(rows):
                            my_tree.move(row, my_tree.parent(
                                row), my_tree.index(row)+1)
                            x.append(my_tree.index(row)+1)
                        for parent in my_tree.get_children():
                            Dattta.append(my_tree.item(parent)["values"])
                        for lpo in range(len(Dattta)):
                            try:
                                Dattta[lpo][7] = 'D-'+Dattta[lpo][0][6:len(
                                    Dattta[lpo][0])]+'-' + Dattta[lpo][0][3:5]+'-'+str(lpo+1)
                            except:
                                donothing = 0
                        for parent in my_tree.get_children():
                            my_tree.delete(parent)
                        for i in Dattta:
                            mytag = 'normal'
                            if i[2] == 'CARTE BANCAIRE':
                                mytag = 'blue'
                            if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                mytag = 'red'
                            my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                            mytag = 'normal'
                        for i in x:
                            if children:
                                my_tree.selection_add(
                                    my_tree.get_children()[i-1])
                                my_tree.focus_set()
                                my_tree.focus(my_tree.get_children()[i-1])
                    ope = 0
                    return

                def idin():
                    global ope
                    if ope == 0:
                        ope = 1
                        curItem = my_tree.focus()
                        values = my_tree.item(curItem, "values")
                        if len(my_tree.selection()) < 2:
                            if len(curItem) > 0:
                                TopIdin = Toplevel()
                                TopIdin.title("Modification de l'ID")
                                TopIdin.geometry("300x150")
                                icon = PhotoImage(file='logo-light.png')
                                window.tk.call(
                                    'wm', 'iconphoto', TopIdin._w, icon)
                                TopIdin.resizable(width=0, height=0)
                                l60 = Label(TopIdin, text="Merci de saisir le nouveau numéro \nd'attestation pour ce paiement.",
                                            font=('Times', 11, 'bold'))
                                l61 = Label(TopIdin, text="D-"+values[0][6:len(values[0])]+'-'+values[0][3:5]+'-',
                                            font=('Times', 11, 'bold'))
                                e60 = Entry(
                                    TopIdin, textvariable=Date, width=5)
                                l60.pack(pady=10)
                                e60.pack()
                                l61.place(x=64, y=58)
                                e60.focus()

                                def hantaTchouf():
                                    global ope
                                    ope = 0
                                    e60.delete(0, END)
                                    TopIdin.destroy()

                                def changeID():
                                    global ope
                                    j = 0
                                    sup = Dattta[len(
                                        Dattta)-1][7][8:len(Dattta[len(Dattta)-1][7])]
                                    inf = Dattta[0][7][8:len(Dattta[0][7])]
                                    e600 = e60.get()
                                    try:
                                        if int(e600) <= int(sup) and int(e600) >= int(inf):
                                            try:
                                                for li in range(len(Dattta)):
                                                    if j == 1:
                                                        break
                                                    if Dattta[li][7][8: len(Dattta[li][7])] == e60.get():
                                                        j += 1
                                                        try:
                                                            x = [Dattta[li][0], Dattta[li][1], Dattta[li][2], Dattta[li][3],
                                                                 Dattta[li][4], Dattta[li][5], Dattta[li][6], Dattta[li][7]]
                                                            del Dattta[li]
                                                            Dattta.insert(
                                                                li, [values[0], values[1], values[2], values[3], values[4], values[5], values[6], values[7][0:8]+e60.get()])
                                                        except:
                                                            donothing = 0
                                                for z in range(len(Dattta)):
                                                    if values[7] == Dattta[z][7]:
                                                        del Dattta[z]
                                                        Dattta.insert(
                                                            z, [x[0], x[1], x[2], x[3], x[4], x[5], x[6], values[7]])
                                                        break

                                            except:
                                                donothing = 0
                                            if j != 0:
                                                for parent in my_tree.get_children():
                                                    my_tree.delete(parent)
                                                if reloulo == 0:
                                                    for i in Dattta:
                                                        mytag = 'normal'
                                                        if i[2] == 'CARTE BANCAIRE':
                                                            mytag = 'blue'
                                                        if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                            mytag = 'red'
                                                        my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                            i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                        mytag = 'normal'
                                                    for u in my_tree.get_children():
                                                        if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                            my_tree.focus(u)
                                                            my_tree.selection_set(
                                                                u)
                                                    hantaTchouf()
                                                else:
                                                    for parent in my_tree.get_children():
                                                        my_tree.delete(parent)
                                                    try:
                                                        if bach == 'date':
                                                            for i in Dattta:
                                                                if kn9lb3la in i[0]:
                                                                    mytag = 'normal'
                                                                    if i[2] == 'CARTE BANCAIRE':
                                                                        mytag = 'blue'
                                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                                        mytag = 'red'
                                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                                    mytag = 'normal'
                                                            for u in my_tree.get_children():
                                                                if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                                    my_tree.focus(
                                                                        u)
                                                                    my_tree.selection_set(
                                                                        u)
                                                            hantaTchouf()
                                                        if bach == 'mois':
                                                            for i in Dattta:
                                                                if kn9lb3la in i[1]:
                                                                    mytag = 'normal'
                                                                    if i[2] == 'CARTE BANCAIRE':
                                                                        mytag = 'blue'
                                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                                        mytag = 'red'
                                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                                    mytag = 'normal'
                                                            for u in my_tree.get_children():
                                                                if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                                    my_tree.focus(
                                                                        u)
                                                                    my_tree.selection_set(
                                                                        u)
                                                            hantaTchouf()
                                                        if bach == 'type':
                                                            for i in Dattta:
                                                                if kn9lb3la in i[2]:
                                                                    mytag = 'normal'
                                                                    if i[2] == 'CARTE BANCAIRE':
                                                                        mytag = 'blue'
                                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                                        mytag = 'red'
                                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                                    mytag = 'normal'
                                                            for u in my_tree.get_children():
                                                                if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                                    my_tree.focus(
                                                                        u)
                                                                    my_tree.selection_set(
                                                                        u)
                                                            hantaTchouf()
                                                        if bach == 'nom':
                                                            for i in Dattta:
                                                                if kn9lb3la in i[3]:
                                                                    mytag = 'normal'
                                                                    if i[2] == 'CARTE BANCAIRE':
                                                                        mytag = 'blue'
                                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                                        mytag = 'red'
                                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                                    mytag = 'normal'
                                                            for u in my_tree.get_children():
                                                                if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                                    my_tree.focus(
                                                                        u)
                                                                    my_tree.selection_set(
                                                                        u)
                                                            hantaTchouf()
                                                        if bach == 'montant':
                                                            for i in Dattta:
                                                                if kn9lb3la in i[4]:
                                                                    mytag = 'normal'
                                                                    if i[2] == 'CARTE BANCAIRE':
                                                                        mytag = 'blue'
                                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                                        mytag = 'red'
                                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                                    mytag = 'normal'
                                                            for u in my_tree.get_children():
                                                                if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                                    my_tree.focus(
                                                                        u)
                                                                    my_tree.selection_set(
                                                                        u)
                                                            hantaTchouf()
                                                        if bach == 'détail':
                                                            for i in Dattta:
                                                                if kn9lb3la in i[5]:
                                                                    mytag = 'normal'
                                                                    if i[2] == 'CARTE BANCAIRE':
                                                                        mytag = 'blue'
                                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                                        mytag = 'red'
                                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                                    mytag = 'normal'
                                                            for u in my_tree.get_children():
                                                                if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                                    my_tree.focus(
                                                                        u)
                                                                    my_tree.selection_set(
                                                                        u)
                                                            hantaTchouf()
                                                        if bach == 'mol':
                                                            for i in Dattta:
                                                                if kn9lb3la in i[6]:
                                                                    mytag = 'normal'
                                                                    if i[2] == 'CARTE BANCAIRE':
                                                                        mytag = 'blue'
                                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                                        mytag = 'red'
                                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                                    mytag = 'normal'
                                                            for u in my_tree.get_children():
                                                                if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                                    my_tree.focus(
                                                                        u)
                                                                    my_tree.selection_set(
                                                                        u)
                                                            hantaTchouf()
                                                        if bach == 'att':
                                                            for i in Dattta:
                                                                if kn9lb3la in i[7]:
                                                                    mytag = 'normal'
                                                                    if i[2] == 'CARTE BANCAIRE':
                                                                        mytag = 'blue'
                                                                    if i[2] != 'CARTE BANCAIRE' and i[3] == '':
                                                                        mytag = 'red'
                                                                    my_tree.insert(parent='', index='end', iid=i, text=i, values=(
                                                                        i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=(mytag))
                                                                    mytag = 'normal'
                                                            for u in my_tree.get_children():
                                                                if my_tree.item(u)["values"][7][8:len(my_tree.item(u)["values"][7])] == e60.get():
                                                                    my_tree.focus(
                                                                        u)
                                                                    my_tree.selection_set(
                                                                        u)
                                                            hantaTchouf()
                                                    except:
                                                        donothing = 0
                                        else:
                                            messagebox.showinfo(
                                                title='Invalide !!', message='Merci de saisir un numéro entre '+Dattta[0][7][8:len(Dattta[0][7])]+' et '+Dattta[len(Dattta)-1][7][8:len(Dattta[len(Dattta)-1][7])], parent=TopIdin)
                                            ope = 1
                                            e60.delete(0, END)
                                    except:
                                        messagebox.showinfo(
                                            title='Invalide !!', message='Merci de saisir un numéro entre '+Dattta[0][7][8:len(Dattta[0][7])]+' et '+Dattta[len(Dattta)-1][7][8:len(Dattta[len(Dattta)-1][7])], parent=TopIdin)
                                        ope = 1
                                        e60.delete(0, END)

                                def hantaTchouf():
                                    global ope
                                    ope = 0
                                    e60.delete(0, END)
                                    TopIdin.destroy()
                                submitbutton = Button(
                                    TopIdin, text='Enregistrer', command=lambda: changeID())
                                submitbutton.configure(
                                    font=('Times', 11, 'bold'), bg='green', fg='white')
                                submitbutton.place(x=150, y=100)
                                cancelbutton = Button(
                                    TopIdin, text='Annulé', command=lambda: hantaTchouf())
                                cancelbutton.configure(
                                    font=('Times', 11, 'bold'), bg='red', fg='white')
                                cancelbutton.place(x=60, y=100)
                                TopIdin.protocol(
                                    "WM_DELETE_WINDOW", hantaTchouf)
                                TopIdin.bind(
                                    "<Escape>", lambda e: hantaTchouf())
                                TopIdin.bind(
                                    "<Return>", lambda e: changeID())
                            else:
                                messagebox.showinfo(
                                    title='!!', message='Merci de selectionner une ligne')
                                ope = 0
                        else:
                            messagebox.showinfo(
                                title='!!', message='Merci de selectionner une seule ligne')
                            ope = 0
                    return

                def telechCSV(tree):
                    global laDate
                    global Dattta
                    global année
                    global télo
                    x = laDate
                    laDate = 'D-' + \
                        str(laDate[6:len(laDate)])+'-'+str(laDate[3:5])
                    try:
                        wb = load_workbook(str(laDate)+".xlsx", read_only=True)
                        if laDate in wb.sheetnames:
                            télo = 1
                    except:
                        télo = 0
                    excel = xlsxwriter.Workbook(str(laDate)+".xlsx")
                    fiche = excel.add_worksheet(laDate)
                    fiche.autofilter('A1:H11')
                    format1 = excel.add_format()
                    format1.set_bg_color('#00B0F0')
                    format1.set_border()
                    format1.set_border_color('#000000')
                    format1.set_bold()
                    format1.set_center_across('center_across')
                    format1.set_shrink()
                    format1.set_font_color('#44546A')
                    format1.set_font_size(10)

                    format2 = excel.add_format()
                    format2.set_bg_color('#FFFFFF')
                    format2.set_border()
                    format2.set_border_color('#000000')
                    format2.set_center_across('center_across')
                    format2.set_shrink()
                    format2.set_font_size(9)

                    format3 = excel.add_format()
                    format3.set_bg_color('#FFFFFF')
                    format3.set_border()
                    format3.set_border_color('#000000')
                    format3.set_align('right')
                    format3.set_shrink()
                    format3.set_font_size(9)

                    fiche.write(0, 0, 'Dates', format1)
                    fiche.write(0, 1, 'Mois', format1)
                    fiche.write(0, 2, 'Type', format1)
                    fiche.write(0, 3, "NOM Donneur d'ordre", format1)
                    fiche.write(0, 4, 'Montant', format1)
                    fiche.write(0, 5, 'Détail', format1)
                    fiche.write(0, 6, 'Montant en lettre', format1)
                    fiche.write(0, 7, 'N° Attestation', format1)
                    if len(Dattta) > 1:
                        for i in range(len(Dattta)):
                            try:
                                if Dattta[i][7][8:len(Dattta[i][7])] == '1':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'01'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '2':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'02'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '3':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'03'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '4':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'04'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '5':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'05'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '6':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'06'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '7':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'07'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '8':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'08'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '9':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'09'
                                if Dattta[i][3] == '':
                                    Dattta[i][3] = ' '
                                if Dattta[i][2] == '':
                                    Dattta[i][2] = ' '
                                if Dattta[i][1] == '':
                                    Dattta[i][1] = ' '
                                if Dattta[i][0] == '':
                                    Dattta[i][0] = ' '
                                if Dattta[i][4] == '':
                                    Dattta[i][4] = ' '
                                if Dattta[i][5] == '':
                                    Dattta[i][5] = ' '
                                if Dattta[i][6] == '':
                                    Dattta[i][6] = ' '
                                if Dattta[i][7] == '':
                                    Dattta[i][7] = ' '
                                Dattta[i][0] = Dattta[i][0][0:6]+"20"+str(année)[2:len(str(année))]
                            except:
                                donothing = 0
                            fiche.write(i+1, 0, Dattta[i][0], format2)
                            fiche.write(i+1, 1, Dattta[i][1], format2)
                            fiche.write(i+1, 2, Dattta[i][2], format2)
                            fiche.write(i+1, 3, Dattta[i][3], format2)
                            fiche.write(i+1, 4, Dattta[i][4], format3)
                            fiche.write(i+1, 5, Dattta[i][5], format2)
                            fiche.write(i+1, 6, Dattta[i][6], format2)
                            fiche.write(i+1, 7, Dattta[i][7], format2)
                            fiche.set_column(0, 0, 13)
                            fiche.set_column(1, 1, 13)
                            fiche.set_column(2, 2, 22)
                            fiche.set_column(3, 3, 32)
                            fiche.set_column(4, 4, 12)
                            fiche.set_column(5, 5, 12)
                            fiche.set_column(6, 6, 60)
                            fiche.set_column(7, 7, 15)
                            fiche.set_row(i, 12)
                        excel.close()
                        for i in Dattta:
                            i[0] = i[0][0:5]+'/'+i[0][8:len(i[0])]
                        for i in range(len(Dattta)):
                            try:
                                if Dattta[i][3] == ' ':
                                    Dattta[i][3] = ''
                                if Dattta[i][2] == ' ':
                                    Dattta[i][2] = ''
                                if Dattta[i][1] == ' ':
                                    Dattta[i][1] = ''
                                if Dattta[i][0] == ' ':
                                    Dattta[i][0] = ''
                                if Dattta[i][4] == ' ':
                                    Dattta[i][4] = ''
                                if Dattta[i][5] == ' ':
                                    Dattta[i][5] = ''
                                if Dattta[i][6] == ' ':
                                    Dattta[i][6] = ''
                                if Dattta[i][7] == ' ':
                                    Dattta[i][7] = ''
                                if Dattta[i][7][8:len(Dattta[i][7])] == '01':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'1'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '02':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'2'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '03':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'3'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '04':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'4'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '05':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'5'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '06':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'6'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '07':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'7'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '08':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'8'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '09':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'9'
                            except:
                                donothing = 0
                        if télo == 1:
                            messagebox.showinfo(
                                title='fichier XLSX enregistré', message='Fichier Excel "'+str(laDate)+'.xlsx" mis à jour.')
                        else:
                            messagebox.showinfo(
                                title='fichier XLSX téléchargé', message='Fichier Excel "'+str(laDate)+'.xlsx" créé.')
                        laDate = x
                    else:
                        messagebox.showinfo(
                            title='!!', message='NO DATA !!')
                    return

                def att(tree):
                    citrop=0
                    if len(tree.selection()) < 1:
                        messagebox.showinfo(
                            title='!!', message='Merci de selectionner une ou plusieur lignes')
                    elif len(tree.selection()) == len(Dattta):
                        for i in range(len(Dattta)):
                            try:
                                if Dattta[i][7][8:len(Dattta[i][7])] == '1':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'01'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '2':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'02'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '3':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'03'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '4':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'04'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '5':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'05'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '6':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'06'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '7':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'07'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '8':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'08'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '9':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'09'
                                if Dattta[i][3] == '':
                                    Dattta[i][3] = ' '
                                if Dattta[i][2] == '':
                                    Dattta[i][2] = ' '
                                if Dattta[i][1] == '':
                                    Dattta[i][1] = ' '
                                if Dattta[i][0] == '':
                                    Dattta[i][0] = ' '
                                if Dattta[i][4] == '':
                                    Dattta[i][4] = ' '
                                if Dattta[i][5] == '':
                                    Dattta[i][5] = ' '
                                if Dattta[i][6] == '':
                                    Dattta[i][6] = ' '
                                if Dattta[i][7] == '':
                                    Dattta[i][7] = ' '
                                
                            except:
                                donotin = 0
                        for i in range(len(Dattta)):
                            D_22_09_10 = Dattta[i][7]
                            Fait_le = Dattta[i][0]
                            Montant = Dattta[i][4]
                            Montant_en_lettre = Dattta[i][6]
                            Nom = Dattta[i][3]
                            Type = Dattta[i][2]
                            mois=''
                            if Fait_le[3:5] == '01':
                                mois = 'Janvier'
                            if Fait_le[3:5] == '02':
                                mois = 'Février'
                            if Fait_le[3:5] == '03':
                                mois = 'Mars'
                            if Fait_le[3:5] == '04':
                                mois = 'Avril'
                            if Fait_le[3:5] == '05':
                                mois = 'Mai'
                            if Fait_le[3:5] == '06':
                                mois = 'Juin'
                            if Fait_le[3:5] == '07':
                                mois = 'Juillet'
                            if Fait_le[3:5] == '08':
                                mois = 'Août'
                            if Fait_le[3:5] == '09':
                                mois = 'Septembre'
                            if Fait_le[3:5] == '10':
                                mois = 'Octobre'
                            if Fait_le[3:5] == '11':
                                mois = 'Novembre'
                            if Fait_le[3:5] == '12':
                                mois = 'Décembre'
                            pdf = FPDF(orientation='P', format='A4')
                            pdf.add_page()

                            pdf.set_xy(90,65)
                            pdf.set_font("times",size=21,style='BU')
                            pdf.cell(txt='Attestation de Don',w=30,align='C')
                            pdf.image('logo-att.png',80, 10,w = 50,h=40)
                            pdf.set_xy(91.5,75)
                            pdf.cell(txt=D_22_09_10,w=28,align='C')

                            text1="Je soussignée, Mme Bouchra OUTAGHANI, Trésorière Générale de JADARA FOUNDATION, atteste par la présente avoir reçu la somme de "
                            text2=Montant+" dirhams ("+Montant_en_lettre+" dirhams)"
                            if Type.lower()=="espèce" or Type.lower()=="espece":
                                text3=" en "
                            else:
                                text3=" par "
                            text4=Type+" "
                            text5="de "
                            text6=Nom+"."
                            text7="La contribution de "
                            text8=Nom+" "
                            text9="participera au financement de la mission de JADARA Foudation telle que prévue par ses status accessibles sur son site web www.jadara.foundation."
                            text10=Nom+" "
                            text11="peut accéder au rapport morale et financier de JADARA Foudation sur son site web www.jadara.foudation."
                            text12="Jadara Foudation est une association reconnue d'utilité publique par le Décret N°2-13-339 du 29 Avril 2013 tel que modifié par le Décret N°2.22.834 du 19 octobre 2022."
                            text13="Cette attestation est délivrée pour servir et valoir ce que de droit."

                            text14="Fait à Casablanca, le "
                            text15=Fait_le[0:2]+" "+mois+" "+"20"+Fait_le[6:len(Fait_le)]

                            text16="Bouchra OUTAGHANI"

                            pdf.set_auto_page_break("ON", margin = 0.0)
                            pdf.set_font("times",size=12)
                            pdf.set_xy(20,105)
                            pdf.multi_cell(w=170,h=5,txt=text1+"**"+text2+"**"+text3+"**"+text4+"**"+text5+"**"+text6+"**"+"\n\n"+text7+"**"+text8+"**"+text9+"\n\n"+"**"+text10+"**"+text11+"\n\n"+text12+"\n\n"+text13,markdown=True,
                                            align='L')

                            pdf.set_font("times",size=11)
                            pdf.set_xy(100,200)
                            pdf.multi_cell(w=90,h=5,txt=text14+"**"+text15+"**"+"\n\n"+"**"+text16+"**",markdown=True,align='R')
                            pdf.set_xy(100,215)
                            pdf.multi_cell(w=90,h=5,txt="**Trésorière Générale**",markdown=True,align='R')
                            pdf.set_font("times",size=9)
                            pdf.set_xy(100,220)
                            pdf.multi_cell(w=90,h=5,txt="**P.O**",markdown=True,align='R')
                            pdf.set_xy(100,225)
                            pdf.multi_cell(w=90,h=5,txt="**Bochra CHABBOUBA ELIDRISSI**",markdown=True,align='R')
                            pdf.set_xy(100,230)
                            pdf.multi_cell(w=90,h=5,txt="**Responsable Administrative et Financière**",markdown=True,align='R')



                            pdf.set_fill_color(193, 153, 9)
                            pdf.set_xy(8,275)
                            pdf.multi_cell(w=0,h=0.5,txt="",fill=True)

                            pdf.set_text_color(45, 82, 158)
                            pdf.set_font("times",size=14,style="B")
                            pdf.set_xy(8,280)
                            pdf.multi_cell(w=0,h=5,txt="JADARA Foundation")

                            pdf.set_text_color(193, 153, 9)
                            pdf.set_font("times",size=7.5,style="")
                            pdf.set_xy(8,285)
                            pdf.multi_cell(w=0,h=5,txt="Jadara Foundation, association reconnue d'utilité publique selon de décret N°2.22.834")

                            pdf.set_text_color(0, 0, 0)
                            pdf.set_font("times",size=7.5,style="")
                            pdf.set_xy(8,289)
                            pdf.multi_cell(w=0,h=5,txt="**Adresse**: 295, Bd Abdelmoumen C24, angle Rue la Percée, 20100, Casablanca",markdown=True)


                            pdf.set_font("times",size=8,style="")
                            pdf.set_xy(107,279)     
                            pdf.multi_cell(w=40,h=5,txt="**T**: +212 522 861 880\n**F**: +212 522 864 178\n**Mail**: contact@jadara.foundation",markdown=True)

                            pdf.set_font("times",size=8,style="")
                            pdf.set_xy(158,283)     
                            pdf.multi_cell(w=40,h=5,txt="**www.jadara.foundation**",markdown=True)


                            pdf.set_xy(152,275)
                            pdf.multi_cell(w=0.5,h=20,txt="",fill=True)


                            pdf.set_xy(102,275)
                            pdf.multi_cell(w=0.5,h=20,txt="",fill=True)




                            pdf.output(str(D_22_09_10)+".pdf")
                            citrop=1
                        for i in range(len(Dattta)):
                            try:
                                if Dattta[i][3] == ' ':
                                    Dattta[i][3] = ''
                                if Dattta[i][2] == ' ':
                                    Dattta[i][2] = ''
                                if Dattta[i][1] == ' ':
                                    Dattta[i][1] = ''
                                if Dattta[i][0] == ' ':
                                    Dattta[i][0] = ''
                                if Dattta[i][4] == ' ':
                                    Dattta[i][4] = ''
                                if Dattta[i][5] == ' ':
                                    Dattta[i][5] = ''
                                if Dattta[i][6] == ' ':
                                    Dattta[i][6] = ''
                                if Dattta[i][7] == ' ':
                                    Dattta[i][7] = ''
                                if Dattta[i][7][8:len(Dattta[i][7])] == '01':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'1'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '02':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'2'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '03':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'3'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '04':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'4'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '05':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'5'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '06':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'6'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '07':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'7'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '08':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'8'
                                if Dattta[i][7][8:len(Dattta[i][7])] == '09':
                                    Dattta[i][7] = Dattta[i][7][0:8]+'9'
                            except:
                                donothing = 0
                    else:

                        lololo = tree.selection()

                        for i in lololo:
                            j = 0
                            l9it = 0
                            for y in Dattta:
                                if l9it == 1:
                                    break
                                if l9it == 0:
                                    if tree.selection()[tree.selection().index(i)].split()[0] == y[7] or y[7] == tree.selection()[tree.selection().index(i)].split()[len(tree.selection()[tree.selection().index(i)].split())-1]:
                                        try:
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '1':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'01'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '2':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'02'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '3':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'03'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '4':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'04'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '5':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'05'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '6':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'06'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '7':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'07'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '8':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'08'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '9':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'09'
                                            if Dattta[j][3] == '':
                                                Dattta[j][3] = ' '
                                            if Dattta[j][2] == '':
                                                Dattta[j][2] = ' '
                                            if Dattta[j][1] == '':
                                                Dattta[j][1] = ' '
                                            if Dattta[j][0] == '':
                                                Dattta[j][0] = ' '
                                            if Dattta[j][4] == '':
                                                Dattta[j][4] = ' '
                                            if Dattta[j][5] == '':
                                                Dattta[j][5] = ' '
                                            if Dattta[j][6] == '':
                                                Dattta[j][6] = ' '
                                            if Dattta[j][7] == '':
                                                Dattta[j][7] = ' '
                                        except:
                                            donotin = 0
                                        D_22_09_10 = Dattta[j][7]
                                        Fait_le = Dattta[j][0]
                                        Montant = Dattta[j][4]
                                        Montant_en_lettre = Dattta[j][6]
                                        Nom = Dattta[j][3]
                                        Type = Dattta[j][2]
                                        mois=''
                                        if Fait_le[3:5] == '01':
                                            mois = 'Janvier'
                                        if Fait_le[3:5] == '02':
                                            mois = 'Février'
                                        if Fait_le[3:5] == '03':
                                            mois = 'Mars'
                                        if Fait_le[3:5] == '04':
                                            mois = 'Avril'
                                        if Fait_le[3:5] == '05':
                                            mois = 'Mai'
                                        if Fait_le[3:5] == '06':
                                            mois = 'Juin'
                                        if Fait_le[3:5] == '07':
                                            mois = 'Juillet'
                                        if Fait_le[3:5] == '08':
                                            mois = 'Août'
                                        if Fait_le[3:5] == '09':
                                            mois = 'Septembre'
                                        if Fait_le[3:5] == '10':
                                            mois = 'Octobre'
                                        if Fait_le[3:5] == '11':
                                            mois = 'Novembre'
                                        if Fait_le[3:5] == '12':
                                            mois = 'Décembre'
                                        pdf = FPDF(orientation='P', format='A4')
                                        pdf.add_page()
                                        pdf.set_xy(90,65)
                                        pdf.set_font("times",size=21,style='BU')
                                        pdf.cell(txt='Attestation de Don',w=30,align='C')
                                        pdf.image('logo-att.png',80, 10,w = 50,h=40)
                                        pdf.set_xy(91.5,75)
                                        pdf.cell(txt=D_22_09_10,w=28,align='C')

                                        text1="Je soussignée, Mme Bouchra OUTAGHANI, Trésorière Générale de JADARA FOUNDATION, atteste par la présente avoir reçu la somme de "
                                        text2=Montant+" dirhams ("+Montant_en_lettre+" dirhams)"
                                        if Type.lower()=="espèce" or Type.lower()=="espece":
                                            text3=" en "
                                        else:
                                            text3=" par "
                                        text4=Type+" "
                                        text5="de "
                                        text6=Nom+"."
                                        text7="La contribution de "
                                        text8=Nom+" "
                                        text9="participera au financement de la mission de JADARA Foudation telle que prévue par ses status accessibles sur son site web www.jadara.foundation."
                                        text10=Nom+" "
                                        text11="peut accéder au rapport morale et financier de JADARA Foudation sur son site web www.jadara.foudation."
                                        text12="Jadara Foudation est une association reconnue d'utilité publique par le Décret N°2-13-339 du 29 Avril 2013 tel que modifié par le Décret N°2.22.834 du 19 octobre 2022."
                                        text13="Cette attestation est délivrée pour servir et valoir ce que de droit."

                                        text14="Fait à Casablanca, le "
                                        text15=Fait_le[0:2]+" "+mois+" "+"20"+Fait_le[6:len(Fait_le)]
                                        text16="Bouchra OUTAGHANI"

                                        pdf.set_auto_page_break("ON", margin = 0.0)
                                        pdf.set_font("times",size=12)
                                        pdf.set_xy(20,105)
                                        pdf.multi_cell(w=170,h=5,txt=text1+"**"+text2+"**"+text3+"**"+text4+"**"+text5+"**"+text6+"**"+"\n\n"+text7+"**"+text8+"**"+text9+"\n\n"+"**"+text10+"**"+text11+"\n\n"+text12+"\n\n"+text13,markdown=True,
                                                        align='L')

                                        pdf.set_font("times",size=11)
                                        pdf.set_xy(100,200)
                                        pdf.multi_cell(w=90,h=5,txt=text14+"**"+text15+"**"+"\n\n"+"**"+text16+"**",markdown=True,align='R')
                                        pdf.set_xy(100,215)
                                        pdf.multi_cell(w=90,h=5,txt="**Trésorière Générale**",markdown=True,align='R')
                                        pdf.set_font("times",size=9)
                                        pdf.set_xy(100,220)
                                        pdf.multi_cell(w=90,h=5,txt="**P.O**",markdown=True,align='R')
                                        pdf.set_xy(100,225)
                                        pdf.multi_cell(w=90,h=5,txt="**Bochra CHABBOUBA ELIDRISSI**",markdown=True,align='R')
                                        pdf.set_xy(100,230)
                                        pdf.multi_cell(w=90,h=5,txt="**Responsable Administrative et Financière**",markdown=True,align='R')

                                        pdf.set_fill_color(193, 153, 9)
                                        pdf.set_xy(8,275)
                                        pdf.multi_cell(w=0,h=0.5,txt="",fill=True)

                                        pdf.set_text_color(45, 82, 158)
                                        pdf.set_font("times",size=14,style="B")
                                        pdf.set_xy(8,280)
                                        pdf.multi_cell(w=0,h=5,txt="JADARA Foundation")

                                        pdf.set_text_color(193, 153, 9)
                                        pdf.set_font("times",size=7.5,style="")
                                        pdf.set_xy(8,285)
                                        pdf.multi_cell(w=0,h=5,txt="Jadara Foundation, association reconnue d'utilité publique selon de décret N°2.22.834")

                                        pdf.set_text_color(0, 0, 0)
                                        pdf.set_font("times",size=7.5,style="")
                                        pdf.set_xy(8,289)
                                        pdf.multi_cell(w=0,h=5,txt="**Adresse**: 295, Bd Abdelmoumen C24, angle Rue la Percée, 20100, Casablanca",markdown=True)

                                        pdf.set_font("times",size=8,style="")
                                        pdf.set_xy(107,279)     
                                        pdf.multi_cell(w=40,h=5,txt="**T**: +212 522 861 880\n**F**: +212 522 864 178\n**Mail**: contact@jadara.foundation",markdown=True)

                                        pdf.set_font("times",size=8,style="")
                                        pdf.set_xy(158,283)     
                                        pdf.multi_cell(w=40,h=5,txt="**www.jadara.foundation**",markdown=True)

                                        pdf.set_xy(152,275)
                                        pdf.multi_cell(w=0.5,h=20,txt="",fill=True)

                                        pdf.set_xy(102,275)
                                        pdf.multi_cell(w=0.5,h=20,txt="",fill=True)

                                        pdf.output(str(D_22_09_10)+".pdf")
                                        citrop=1

                                        try:
                                            if Dattta[j][3] == ' ':
                                                Dattta[j][3] = ''
                                            if Dattta[j][2] == ' ':
                                                Dattta[j][2] = ''
                                            if Dattta[j][1] == ' ':
                                                Dattta[j][1] = ''
                                            if Dattta[j][0] == ' ':
                                                Dattta[j][0] = ''
                                            if Dattta[j][4] == ' ':
                                                Dattta[j][4] = ''
                                            if Dattta[j][5] == ' ':
                                                Dattta[j][5] = ''
                                            if Dattta[j][6] == ' ':
                                                Dattta[j][6] = ''
                                            if Dattta[j][7] == ' ':
                                                Dattta[j][7] = ''
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '01':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'1'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '02':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'2'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '03':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'3'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '04':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'4'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '05':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'5'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '06':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'6'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '07':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'7'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '08':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'8'
                                            if Dattta[j][7][8:len(Dattta[j][7])] == '09':
                                                Dattta[j][7] = Dattta[j][7][0:8]+'9'
                                        except:
                                            donothing = 0
                                        l9it = 1
                                j += 1
                    if citrop==1:
                        messagebox.showinfo(
                            title='', message="Fichier(s) enregistré(s).", parent=window)
                    return
                rowscolor = Button(
                    frame4, text='', background='gray', command=changerowcolor)
                rowscolor.place(x=75, y=5)

                selall = Button(frame5, text='}', command=selectibabahom)
                selall.place(x=2, y=25)
                moveup = Button(frame5, text='⤊', command=up)
                moveup.place(x=2, y=130)
                movedown = Button(frame5, text='⤋', command=down)
                movedown.place(x=2, y=194)
                idinput = Button(frame5, text='•', command=idin)
                idinput.place(x=3, y=162)
                selall1 = Button(frame5, text='}', command=selectibabahom)
                selall1.place(x=2, y=295)

                télécharger = Button(
                    frame1, text='Télécharger xlsx', command=lambda: telechCSV(my_tree))
                télécharger.pack(pady=20)
                lo8 = Label(frame1, text='Ctrl-X', fg='white',
                            background='black', font=('Times', 10))
                lo8.pack()

                rechercher = Button(frame3, text='Rechercher',
                                    command=lambda: menuRecherch())
                delete = Button(frame3, text='Supprimer',
                                command=lambda: deleteData(my_tree))
                add = Button(frame3, text='Ajouter',
                             command=lambda: addData(my_tree))
                edite = Button(frame3, text='Modifier',
                               command=lambda: editData(my_tree))
                Attestation = Button(
                    frame3, text='  Attestation(s)', command=lambda: att(my_tree))

                window.bind('<Control-k>', lambda *args: selectibabahom())
                window.bind('<Control-K>', lambda *args: selectibabahom())
                window.bind("<Control-s>", lambda e: deleteData(my_tree))
                window.bind("<Control-f>", lambda e: menuRecherch())
                window.bind("<Control-F>", lambda e: menuRecherch())
                window.bind("<Control-a>", lambda e: addData(my_tree))
                window.bind("<Control-P>", lambda e: showPDF())
                window.bind("<Control-p>", lambda e: showPDF())
                window.bind("<Control-Z>", lambda e: rj3Data(my_tree))
                window.bind("<Control-z>", lambda e: rj3Data(my_tree))
                window.bind("<Control-m>", lambda e: editData(my_tree))
                my_tree.bind("<Double-1>", lambda e: editData(my_tree))
                window.bind("<Control-S>", lambda e: deleteData(my_tree))
                window.bind("<Control-A>", lambda e: addData(my_tree))
                window.bind("<Control-M>", lambda e: editData(my_tree))
                window.bind("+", lambda e: up())
                window.bind("=", lambda e: down())
                window.bind("<Control-I>", lambda e: idin())
                window.bind("<Control-i>", lambda e: idin())
                window.bind("<Control-T>", lambda e: att(my_tree))
                window.bind("<Control-t>", lambda e: att(my_tree))
                window.bind("<Control-x>", lambda e: telechCSV(my_tree))
                window.bind("<Control-X>", lambda e: telechCSV(my_tree))
                rechercher.pack(pady=20, padx=20)
                lo1 = Label(frame3, text='Ctrl-F', fg='white',
                            background='black', font=('Times', 10))
                lo1.pack()
                delete.pack(pady=20, padx=20)
                lo2 = Label(frame3, text='Ctrl-S', fg='white',
                            background='black', font=('Times', 10))
                lo2.pack()
                add.pack(pady=20, padx=20)
                lo3 = Label(frame3, text='Ctrl-A', fg='white',
                            background='black', font=('Times', 10))
                lo3.pack()
                edite.pack(pady=20, padx=20)
                lo4 = Label(frame3, text='Ctrl-M', fg='white',
                            background='black', font=('Times', 10))
                lo4.pack()
                Attestation.pack(pady=20, padx=20)
                lo5 = Label(frame3, text='Ctrl-T', fg='white',
                            background='black', font=('Times', 10))
                lo5.pack()
                return
            return

        frame1.pack(fill='y', pady=10, padx=10, side='right')
        frame2.pack(fill='both', side='left', expand=True, padx=10, pady=10)
        charger = Button(frame1, text='Charger les données',
                         command=traitement)
        charger.pack(pady=20)
        lo20 = Label(frame1, text='Entrée', fg='white',
                     background='black', font=('Times', 10))
        lo20.pack()
        init = Button(frame1, text='Rénitialiser les données', command=reset)
        lo12 = Label(frame1, text='Ctrl-R', fg='white',
                     background='black', font=('Times', 10))
        lo12.pack(side='bottom')
        init.pack(side='bottom', pady=20, padx=5)
        l = Label(frame2, text=os.path.basename(filename))
        l.pack(pady=5, padx=5)
        window.bind("<Return>", lambda e: traitement())
        window.bind("<Control-r>", lambda e: reset())
        window.bind("<Control-R>", lambda e: reset())

    b = Button(text="Attestation", command=lambda: secondaire())
    a = Button(text="Excel", command=lambda: principal())
    window.bind("<Control-T>", lambda e: secondaire())
    window.bind("<Control-t>", lambda e: secondaire())
    a.place(relx=0.5, rely=0.3, anchor=CENTER)
    lo14 = Label(text='Entrée', fg='white',
                 background='black', font=('Times', 10))
    lo14.place(relx=0.5, rely=0.4, anchor=CENTER)
    b.place(relx=0.5, rely=0.7, anchor=CENTER)
    lo13 = Label(text='Ctrl-T', fg='white',
                 background='black', font=('Times', 10))
    lo13.place(relx=0.5, rely=0.8, anchor=CENTER)
    window.title(" JADARA ")
    window.bind("<Return>", lambda e: principal())


window.mainloop()