from lxml import etree
from openpyxl import load_workbook
import pandas as pd
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
file_path = filedialog.askdirectory()

treea = etree.parse(file_path+"/NACINI_UPORABE.gml")
treeb = etree.parse(file_path+"/CESTICE.gml")
treec = etree.parse(file_path+"/POSJEDOVNI_LISTOVI.gml")
treed = etree.parse(file_path+"/UPISANE_OSOBE.gml")
treee = etree.parse(file_path+"/CESTICE_ZGRADE.gml")
treef = etree.parse(file_path+"/ZGRADE.gml")
treeg= etree.parse(file_path+"/ADRESE_UPISANIH_OSOBA.gml")

treeazk = etree.parse(file_path+"/ZK_CESTICE.gml")
treebzk = etree.parse(file_path+"/ULOSCI.gml")
treeczk = etree.parse(file_path+"/ZK_NACINI_UPORABE.gml")
treezzk = etree.parse(file_path+"/ZK_ADRESE_CESTICE.gml")
treedzk = etree.parse(file_path+"/ZK_VLASNICI.gml")
treeezk = etree.parse(file_path+"/ZK_ADRESE_VLASNIKA.gml")
treeaz = etree.parse(file_path+"/ADRESE_ZGRADE.gml")
treeax = etree.parse(file_path+"/ZK_ZGRADE.gml")

rootazk = treeazk.getroot()
rootbzk = treebzk.getroot()
rootczk = treeczk.getroot()
rootdzk = treedzk.getroot()
rootezk = treeezk.getroot()
rootzzk = treezzk.getroot()
rootaz = treeaz.getroot()
rootax = treeax.getroot()

roota = treea.getroot()
rootb = treeb.getroot()
rootc = treec.getroot()
rootd = treed.getroot()
roote = treee.getroot()
rootf = treef.getroot()
rootg = treeg.getroot()

lista_posjedovni_list_broj_nacini_uporabe = []
for POSJEDOVNI_LIST_BROJ in roota.iter("{http://www.uredjenazemlja.hr}POSJEDOVNI_LIST_BROJ"):
    lista_posjedovni_list_broj_nacini_uporabe.append(POSJEDOVNI_LIST_BROJ.text)

lista_povrsina_nacini_uporabe = []
for POVRSINA in roota.iter("{http://www.uredjenazemlja.hr}POVRSINA"):
    lista_povrsina_nacini_uporabe.append(POVRSINA.text)

lista_cestica_id_nacini_uporabe = []
for CESTICA_ID in roota.iter("{http://www.uredjenazemlja.hr}CESTICA_ID"):
    lista_cestica_id_nacini_uporabe.append(CESTICA_ID.text)

lista_naziv_vrste_uporabe_nacini_uporabe = []
for NAZIV_VRSTE_UPORABE in roota.iter("{http://www.uredjenazemlja.hr}NAZIV_VRSTE_UPORABE"):
    lista_naziv_vrste_uporabe_nacini_uporabe.append(NAZIV_VRSTE_UPORABE.text)

lista_cestica_id_cestice = []
for CESTICA_ID in rootb.iter("{http://www.uredjenazemlja.hr}CESTICA_ID"):
    lista_cestica_id_cestice.append(CESTICA_ID.text)

lista_broj_cestice_cestice = []
for BROJ_CESTICE in rootb.iter("{http://www.uredjenazemlja.hr}BROJ_CESTICE"):
    lista_broj_cestice_cestice.append(BROJ_CESTICE.text)

lista_povrsina_graficka_cestice = []
for POVRSINA_GRAFICKA in rootb.iter("{http://www.uredjenazemlja.hr}POVRSINA_GRAFICKA"):
    lista_povrsina_graficka_cestice.append(POVRSINA_GRAFICKA.text)

lista_povrsina_atributna_cestice = []
for POVRSINA_ATRIBUTNA in rootb.iter("{http://www.uredjenazemlja.hr}POVRSINA_ATRIBUTNA"):
    lista_povrsina_atributna_cestice.append(POVRSINA_ATRIBUTNA.text)

lista_adresa_opisna_cestice = []
for ADRESA_OPISNA in rootb.iter("{http://www.uredjenazemlja.hr}ADRESA_OPISNA"):
    lista_adresa_opisna_cestice.append(ADRESA_OPISNA.text)

lista_posjedovni_list_id_posjedovni_listovi = []
for POSJEDOVNI_LIST_ID in rootc.iter("{http://www.uredjenazemlja.hr}POSJEDOVNI_LIST_ID"):
    lista_posjedovni_list_id_posjedovni_listovi.append(POSJEDOVNI_LIST_ID.text)

lista_broj_posjedovni_listovi = []
for BROJ in rootc.iter("{http://www.uredjenazemlja.hr}BROJ"):
    lista_broj_posjedovni_listovi.append(BROJ.text)

lista_posjednik_id_upisane_osobe = []
for POSJEDNIK_ID in rootd.iter("{http://www.uredjenazemlja.hr}POSJEDNIK_ID"):
    lista_posjednik_id_upisane_osobe.append(POSJEDNIK_ID.text)

lista_posjedovni_list_id_upisane_osobe = []
for POSJEDOVNI_LIST_ID in rootd.iter("{http://www.uredjenazemlja.hr}POSJEDOVNI_LIST_ID"):
    lista_posjednik_id_upisane_osobe.append(POSJEDOVNI_LIST_ID.text)

lista_posjednik_id_upisane_osobe = []
for POSJEDNIK_ID in rootd.iter("{http://www.uredjenazemlja.hr}POSJEDNIK_ID"):
    lista_posjednik_id_upisane_osobe.append(POSJEDNIK_ID.text)

lista_posjedovni_list_id_upisane_osobe = []
for POSJEDOVNI_LIST_ID in rootd.iter("{http://www.uredjenazemlja.hr}POSJEDOVNI_LIST_ID"):
    lista_posjedovni_list_id_upisane_osobe.append(POSJEDOVNI_LIST_ID.text)

lista_posjednik_udio_brojnik_upisane_osobe = []
for POSJEDNIK_UDIO_BROJNIK in rootd.iter("{http://www.uredjenazemlja.hr}POSJEDNIK_UDIO_BROJNIK"):
    lista_posjednik_udio_brojnik_upisane_osobe.append(POSJEDNIK_UDIO_BROJNIK.text)

lista_posjednik_udio_nazivnik_upisane_osobe = []
for POSJEDNIK_UDIO_NAZIVNIK in rootd.iter("{http://www.uredjenazemlja.hr}POSJEDNIK_UDIO_NAZIVNIK"):
    lista_posjednik_udio_nazivnik_upisane_osobe.append(POSJEDNIK_UDIO_NAZIVNIK.text)

lista_naziv_upisane_osobe = []
for NAZIV in rootd.iter("{http://www.uredjenazemlja.hr}NAZIV"):
    lista_naziv_upisane_osobe.append(NAZIV.text)

lista_naziv_pravnog_odnosa_upisane_osobe = []
for NAZIV_PRAVNOG_ODNOSA in rootd.iter("{http://www.uredjenazemlja.hr}NAZIV_PRAVNOG_ODNOSA"):
    lista_naziv_pravnog_odnosa_upisane_osobe.append(NAZIV_PRAVNOG_ODNOSA.text)

lista_porezni_broj_upisane_osobe = []
for POREZNI_BROJ in rootd.iter("{http://www.uredjenazemlja.hr}POREZNI_BROJ"):
    lista_porezni_broj_upisane_osobe.append(POREZNI_BROJ.text)

lista_cestica_id_cestice_zgrade = []
for CESTICA_ID in roote.iter("{http://www.uredjenazemlja.hr}CESTICA_ID"):
    lista_cestica_id_cestice_zgrade.append(CESTICA_ID.text)

lista_zgrada_id_cestice_zgrade = []
for ZGRADA_ID in roote.iter("{http://www.uredjenazemlja.hr}ZGRADA_ID"):
    lista_zgrada_id_cestice_zgrade.append(ZGRADA_ID.text)

lista_zgrada_id_zgrade = []
for ZGRADA_ID in rootf.iter("{http://www.uredjenazemlja.hr}ZGRADA_ID"):
    lista_zgrada_id_zgrade.append(ZGRADA_ID.text)

lista_povrsina_zgrade = []
for POVRSINA in rootf.iter("{http://www.uredjenazemlja.hr}POVRSINA"):
    lista_povrsina_zgrade.append(POVRSINA.text)

lista_naziv_vrste_zgrade_zgrade = []
for NAZIV_VRSTE_ZGRADE in rootf.iter("{http://www.uredjenazemlja.hr}NAZIV_VRSTE_ZGRADE"):
    lista_naziv_vrste_zgrade_zgrade.append(NAZIV_VRSTE_ZGRADE.text)

lista_posjedovni_list_broj_zgrade = []
for POSJEDOVNI_LIST_BROJ in rootf.iter("{http://www.uredjenazemlja.hr}POSJEDOVNI_LIST_BROJ"):
    lista_posjedovni_list_broj_zgrade.append(POSJEDOVNI_LIST_BROJ.text)

lista_zk_cestica_id_zk_cestice = []
for ZK_CESTICA_ID in rootazk.iter("{http://www.uredjenazemlja.hr}ZK_CESTICA_ID"):
    lista_zk_cestica_id_zk_cestice.append(ZK_CESTICA_ID.text)

lista_broj_cestice_zk_cestice = []
for BROJ_CESTICE in rootazk.iter("{http://www.uredjenazemlja.hr}BROJ_CESTICE"):
    lista_broj_cestice_zk_cestice.append(BROJ_CESTICE.text)

lista_podbroj_cestice_zk_cestice = []
for PODBROJ_CESTICE in rootazk.iter("{http://www.uredjenazemlja.hr}PODBROJ_CESTICE"):
    lista_podbroj_cestice_zk_cestice.append(PODBROJ_CESTICE.text)

lista_ulozak_id_zk_cestice = []
for ULOZAK_ID in rootazk.iter("{http://www.uredjenazemlja.hr}ULOZAK_ID"):
    lista_ulozak_id_zk_cestice.append(ULOZAK_ID.text)

lista_povrsina_zk_cestice = []
for POVRSINA in rootazk.iter("{http://www.uredjenazemlja.hr}POVRSINA"):
    lista_povrsina_zk_cestice.append(POVRSINA.text)

lista_rbr_zk_tijela_zk_cestice = []
for RBR_ZK_TIJELA in rootazk.iter("{http://www.uredjenazemlja.hr}RBR_ZK_TIJELA"):
    lista_rbr_zk_tijela_zk_cestice.append(RBR_ZK_TIJELA.text)

lista_ulozak_id_ulosci = []
for ULOZAK_ID in rootbzk.iter("{http://www.uredjenazemlja.hr}ULOZAK_ID"):
    lista_ulozak_id_ulosci.append(ULOZAK_ID.text)

lista_broj_uloska_ulosci = []
for BROJ_ULOSKA in rootbzk.iter("{http://www.uredjenazemlja.hr}BROJ_ULOSKA"):
    lista_broj_uloska_ulosci.append(BROJ_ULOSKA.text)

lista_zk_cestica_id_zk_nacini_uporabe = []
for ZK_CESTICA_ID in rootczk.iter("{http://www.uredjenazemlja.hr}ZK_CESTICA_ID"):
    lista_zk_cestica_id_zk_nacini_uporabe.append(ZK_CESTICA_ID.text)

lista_zk_naziv_vrste_uporabe_zk_nacini_uporabe = []
for NAZIV_VRSTE_UPORABE in rootczk.iter("{http://www.uredjenazemlja.hr}NAZIV_VRSTE_UPORABE"):
    lista_zk_naziv_vrste_uporabe_zk_nacini_uporabe.append(NAZIV_VRSTE_UPORABE.text)

lista_zk_povrsina_zk_nacini_uporabe = []
for POVRSINA in rootczk.iter("{http://www.uredjenazemlja.hr}POVRSINA"):
    lista_zk_povrsina_zk_nacini_uporabe.append(POVRSINA.text)

lista_adresa_opisna_zk_adrese_cestice = []
for ADRESA_OPISNA in rootzzk.iter("{http://www.uredjenazemlja.hr}ADRESA_OPISNA"):
    lista_adresa_opisna_zk_adrese_cestice.append(ADRESA_OPISNA.text)

lista_zk_cestica_id_zk_adrese_cestice = []
for ZK_CESTICA_ID in rootzzk.iter("{http://www.uredjenazemlja.hr}ZK_CESTICA_ID"):
    lista_zk_cestica_id_zk_adrese_cestice.append(ZK_CESTICA_ID.text)

lista_ulozak_id_zk_vlasnici = []
for ULOZAK_ID in rootdzk.iter("{http://www.uredjenazemlja.hr}ULOZAK_ID"):
    lista_ulozak_id_zk_vlasnici.append(ULOZAK_ID.text)

lista_osoba_upisa_id_zk_vlasnici = []
for OSOBA_UPISA_ID in rootdzk.iter("{http://www.uredjenazemlja.hr}OSOBA_UPISA_ID"):
    lista_osoba_upisa_id_zk_vlasnici.append(OSOBA_UPISA_ID.text)

lista_udio_dijela_brojnik_zk_vlasnici = []
for UDIO_DIJELA_BROJNIK in rootdzk.iter("{http://www.uredjenazemlja.hr}UDIO_DIJELA_BROJNIK"):
    lista_udio_dijela_brojnik_zk_vlasnici.append(UDIO_DIJELA_BROJNIK.text)

lista_udio_dijela_nazivnik_zk_vlasnici = []
for UDIO_DIJELA_NAZIVNIK in rootdzk.iter("{http://www.uredjenazemlja.hr}UDIO_DIJELA_NAZIVNIK"):
    lista_udio_dijela_nazivnik_zk_vlasnici.append(UDIO_DIJELA_NAZIVNIK.text)

lista_naziv_zk_vlasnici = []
for NAZIV in rootdzk.iter("{http://www.uredjenazemlja.hr}NAZIV"):
    lista_naziv_zk_vlasnici.append(NAZIV.text)

lista_ulozak_id_zk_vlasnici = []
for ULOZAK_ID in rootdzk.iter("{http://www.uredjenazemlja.hr}ULOZAK_ID"):
    lista_ulozak_id_zk_vlasnici.append(ULOZAK_ID.text)

lista_porezni_broj_zk_vlasnici = []
for POREZNI_BROJ in rootdzk.iter("{http://www.uredjenazemlja.hr}POREZNI_BROJ"):
    lista_porezni_broj_zk_vlasnici.append(POREZNI_BROJ.text)

lista_udio_brojnik_zk_vlasnici = []
for UDIO_BROJNIK in rootdzk.iter("{http://www.uredjenazemlja.hr}UDIO_BROJNIK"):
    lista_udio_brojnik_zk_vlasnici.append(UDIO_BROJNIK.text)

lista_udio_nazivnik_zk_vlasnici = []
for UDIO_NAZIVNIK in rootdzk.iter("{http://www.uredjenazemlja.hr}UDIO_NAZIVNIK"):
    lista_udio_nazivnik_zk_vlasnici.append(UDIO_NAZIVNIK.text)

lista_rbr_etaze_zk_vlasnici = []
for RBR_ETAZE in rootdzk.iter("{http://www.uredjenazemlja.hr}RBR_ETAZE"):
    lista_rbr_etaze_zk_vlasnici.append(RBR_ETAZE.text)

lista_rbr_zk_tijela_zk_vlasnici = []
for RBR_ZK_TIJELA in rootdzk.iter("{http://www.uredjenazemlja.hr}RBR_ZK_TIJELA"):
    lista_rbr_zk_tijela_zk_vlasnici.append(RBR_ZK_TIJELA.text)

lista_posjednik_id_adrese_upisanih_osoba = []
for POSJEDNIK_ID in rootg.iter("{http://www.uredjenazemlja.hr}POSJEDNIK_ID"):
    lista_posjednik_id_adrese_upisanih_osoba.append(POSJEDNIK_ID.text)

lista_drzava_adrese_upisanih_osoba = []
for DRZAVA in rootg.iter("{http://www.uredjenazemlja.hr}DRZAVA"):
    lista_drzava_adrese_upisanih_osoba.append(DRZAVA.text)

lista_adresa_opisna_adrese_upisanih_osoba = []
for ADRESA_OPISNA in rootg.iter("{http://www.uredjenazemlja.hr}ADRESA_OPISNA"):
    lista_adresa_opisna_adrese_upisanih_osoba.append(ADRESA_OPISNA.text)

lista_postanski_broj_adrese_upisanih_osoba = []
for POSTANSKI_BROJ in rootg.iter("{http://www.uredjenazemlja.hr}POSTANSKI_BROJ"):
    lista_postanski_broj_adrese_upisanih_osoba.append(POSTANSKI_BROJ.text)

lista_naselje_adrese_upisanih_osoba = []
for NASELJE in rootg.iter("{http://www.uredjenazemlja.hr}NASELJE"):
    lista_naselje_adrese_upisanih_osoba.append(NASELJE.text)

lista_ulica_adrese_upisanih_osoba = []
for ULICA in rootg.iter("{http://www.uredjenazemlja.hr}ULICA"):
    lista_ulica_adrese_upisanih_osoba.append(ULICA.text)

lista_kbr_adrese_upisanih_osoba = []
for KBR in rootg.iter("{http://www.uredjenazemlja.hr}KBR"):
    lista_kbr_adrese_upisanih_osoba.append(KBR.text)

lista_osoba_upisa_id_zk_adrese_vlasnika = []
for OSOBA_UPISA_ID in rootezk.iter("{http://www.uredjenazemlja.hr}OSOBA_UPISA_ID"):
    lista_osoba_upisa_id_zk_adrese_vlasnika.append(OSOBA_UPISA_ID.text)

lista_adresa_opisna_zk_adrese_vlasnika = []
for ADRESA_OPISNA in rootezk.iter("{http://www.uredjenazemlja.hr}ADRESA_OPISNA"):
    lista_adresa_opisna_zk_adrese_vlasnika.append(ADRESA_OPISNA.text)

lista_zgrada_id_adrese_zgrade = []
for ZGRADA_ID in rootaz.iter("{http://www.uredjenazemlja.hr}ZGRADA_ID"):
    lista_zgrada_id_adrese_zgrade.append(ZGRADA_ID.text)

lista_adresa_opisna_adrese_zgrade = []
for ADRESA_OPISNA in rootaz.iter("{http://www.uredjenazemlja.hr}ADRESA_OPISNA"):
    lista_adresa_opisna_adrese_zgrade.append(ADRESA_OPISNA.text)

lista_povrsina_zk_zgrade = []
for POVRSINA in rootax.iter("{http://www.uredjenazemlja.hr}POVRSINA"):
    lista_povrsina_zk_zgrade.append(POVRSINA.text)

lista_zk_cestica_id_zk_zgrade = []
for ZK_CESTICA_ID in rootax.iter("{http://www.uredjenazemlja.hr}ZK_CESTICA_ID"):
    lista_zk_cestica_id_zk_zgrade.append(ZK_CESTICA_ID.text)

lista_naziv_vrste_zgrade_zk_zgrade = []
for NAZIV_VRSTE_ZGRADE in rootax.iter("{http://www.uredjenazemlja.hr}NAZIV_VRSTE_ZGRADE"):
    lista_naziv_vrste_zgrade_zk_zgrade.append(NAZIV_VRSTE_ZGRADE.text)

#KREIRANJE DATAFRAMEOVA
####################################################################################
dataa = {"Posjedovni list broj":lista_posjedovni_list_broj_nacini_uporabe,
        "Cestica id":lista_cestica_id_nacini_uporabe,
        "Naziv vrste uporabe":lista_naziv_vrste_uporabe_nacini_uporabe,
        "Povrsina":lista_povrsina_nacini_uporabe}
dataframea = pd.DataFrame(dataa)

datab = {"Cestica id":lista_cestica_id_cestice,
         "Broj cestice":lista_broj_cestice_cestice,
         "Adresa":lista_adresa_opisna_cestice}
dataframeb = pd.DataFrame(datab)

datac = {"Posjedovni list id":lista_posjedovni_list_id_posjedovni_listovi,
        "Posjedovni list broj":lista_broj_posjedovni_listovi}
dataframec = pd.DataFrame(datac)

datad = {"Posjednik id":lista_posjednik_id_upisane_osobe,
         "Posjedovni list id":lista_posjedovni_list_id_upisane_osobe,
         "Posjednik udio brojnik":lista_posjednik_udio_brojnik_upisane_osobe,
         "Posjednik udio nazivnik":lista_posjednik_udio_nazivnik_upisane_osobe,
         "Naziv":lista_naziv_upisane_osobe,
         "Naziv pravnog odnosa":lista_naziv_pravnog_odnosa_upisane_osobe,
         "Porezni broj":lista_porezni_broj_upisane_osobe}
dataframed = pd.DataFrame(datad)
dataframed = dataframed.fillna("")

datag = {"Posjednik id":lista_posjednik_id_adrese_upisanih_osoba,
         "Drzava":lista_drzava_adrese_upisanih_osoba,
         "Adresa opisna":lista_adresa_opisna_adrese_upisanih_osoba,
         "Postanski broj":lista_postanski_broj_adrese_upisanih_osoba,
         "Naselje":lista_naselje_adrese_upisanih_osoba,
         "Ulica":lista_ulica_adrese_upisanih_osoba,
         "Kucni broj":lista_kbr_adrese_upisanih_osoba}
dataframeg = pd.DataFrame(datag)

datae = {"Cestica id":lista_cestica_id_cestice_zgrade,
        "Zgrada id":lista_zgrada_id_cestice_zgrade}
dataframee = pd.DataFrame(datae)

dataf = {"Zgrada id":lista_zgrada_id_zgrade,
         "Posjedovni list broj":lista_posjedovni_list_broj_zgrade,
         "Naziv vrste zgrade":lista_naziv_vrste_zgrade_zgrade,
         "Povrsina":lista_povrsina_zgrade}
dataframef = pd.DataFrame(dataf)
dataframef = dataframef.fillna("")

dataazk = {"Zk cestica id":lista_zk_cestica_id_zk_cestice,
         "Broj cestice brojnik":lista_broj_cestice_zk_cestice,
         "Podbroj cestice":lista_podbroj_cestice_zk_cestice,
         "Ulozak id":lista_ulozak_id_zk_cestice,
         "Povrsina":lista_povrsina_zk_cestice,
         "Redni broj ZK tijela cestice":lista_rbr_zk_tijela_zk_cestice}
dataframeazk = pd.DataFrame(dataazk)

databzk = {"Ulozak id":lista_ulozak_id_ulosci,
         "Broj uloska":lista_broj_uloska_ulosci}
dataframebzk= pd.DataFrame(databzk)

dataczk = {"Zk cestica id":lista_zk_cestica_id_zk_nacini_uporabe,
           "Naziv vrste uporabe":lista_zk_naziv_vrste_uporabe_zk_nacini_uporabe,
           "Povrsina vrste uporabe":lista_zk_povrsina_zk_nacini_uporabe}
dataframeczk = pd.DataFrame(dataczk)

datazzk = {"Zk cestica id":lista_zk_cestica_id_zk_adrese_cestice,
           "Adresa opisna":lista_adresa_opisna_zk_adrese_cestice}
dataframezzk = pd.DataFrame(datazzk)

datadzk = {"Osoba upisa id":lista_osoba_upisa_id_zk_vlasnici,
           "Udio brojnik":lista_udio_brojnik_zk_vlasnici,
           "Udio nazivnik":lista_udio_nazivnik_zk_vlasnici,
           "Udio dijela brojnik":lista_udio_dijela_brojnik_zk_vlasnici,
           "Udio dijela nazivnik":lista_udio_dijela_nazivnik_zk_vlasnici,
           "Redni broj etaze":lista_rbr_etaze_zk_vlasnici,
           "Naziv":lista_naziv_zk_vlasnici,
           "Ulozak id":lista_ulozak_id_zk_vlasnici,
           "Porezni broj":lista_porezni_broj_zk_vlasnici,
           "Redni broj ZK tijela vlasnika":lista_rbr_zk_tijela_zk_vlasnici}
dataframedzk = pd.DataFrame(datadzk)

dataezk = {"Osoba upisa id":lista_osoba_upisa_id_zk_adrese_vlasnika,
           "Adresa":lista_adresa_opisna_zk_adrese_vlasnika}

dataxzk = {"Zk cestica id":lista_zk_cestica_id_zk_zgrade,
           "Povrsina zgrade":lista_povrsina_zk_zgrade,
           "Naziv vrste zgrade":lista_naziv_vrste_zgrade_zk_zgrade}
dataframexzk = pd.DataFrame(dataxzk)
dataframeezk = pd.DataFrame(dataezk)
dataframeezk = dataframeezk.fillna("")

dataaz = {"Zgrada id":lista_zgrada_id_adrese_zgrade,
          "Adresa opisna":lista_adresa_opisna_adrese_zgrade}
dataframeaz = pd.DataFrame(dataaz)


dataframed = dataframed.merge(dataframeg, on="Posjednik id", how="left")
dataframed = dataframed.reset_index()


dataframed = dataframed.merge(dataframec, on="Posjedovni list id", how = "left")

dataframed['Omjer'] = ""
dataframed["Sklopljeno"]=""
dataframed["Adresa"]=""

for i in range(0,len(dataframed["Posjedovni list id"])):
    dataframed.at[i, "Omjer"] = str(dataframed.at[i, "Posjednik udio brojnik"]) + "/" + str(dataframed.at[i, "Posjednik udio nazivnik"])
    if dataframed.at[i,"Adresa opisna"] != None and dataframed.at[i,"Naselje"] == None:
        dataframed.at[i,"Adresa"] = dataframed.at[i,"Adresa opisna"]
    elif dataframed.at[i,"Adresa opisna"] == None and dataframed.at[i,"Naselje"] != None and dataframed.at[i,"Postanski broj"] == None:
        dataframed.at[i, "Adresa"] = dataframed.at[i, "Ulica"]+" "+str(dataframed.at[i, "Kucni broj"]) + ", " + dataframed.at[i, "Naselje"]+", "+dataframed.at[i, "Drzava"]
    elif dataframed.at[i,"Adresa opisna"] == None and dataframed.at[i,"Naselje"] != None:
        dataframed.at[i, "Adresa"] = dataframed.at[i, "Ulica"]+" "+str(dataframed.at[i, "Kucni broj"]) + ", " +dataframed.at[i, "Naselje"]+", "+dataframed.at[i, "Drzava"]
    else:
        pass

for i in range(0,len(dataframed["Posjedovni list id"])):
    if str(dataframed.at[i,"Naziv pravnog odnosa"]) == "VLASNIK" and dataframed.at[i,"Porezni broj"] != "":
        dataframed.at[i,"Sklopljeno"] =dataframed.at[i,"Naziv"]+", "+dataframed.at[i, "Adresa"]+", (Vlasnik), "+"OIB: "+str(dataframed.at[i,"Porezni broj"])
    elif str(dataframed.at[i,"Naziv pravnog odnosa"]) == "VLASNIK" and dataframed.at[i,"Porezni broj"] == "" and dataframed.at[i,"Adresa"] != "":
        dataframed.at[i,"Sklopljeno"] = dataframed.at[i,"Naziv"]+", "+dataframed.at[i, "Adresa"]+", (Vlasnik)"
    elif str(dataframed.at[i,"Naziv pravnog odnosa"]) == "VLASNIK" and dataframed.at[i,"Porezni broj"] == "" and dataframed.at[i,"Adresa"] == "":
        dataframed.at[i,"Sklopljeno"] = dataframed.at[i,"Naziv"]+", (Vlasnik)"
    elif dataframed.at[i,"Porezni broj"] == "":
        dataframed.at[i, "Sklopljeno"] = dataframed.at[i, "Naziv"] + ", " + dataframed.at[i, "Adresa"]
    else:
        dataframed.at[i, "Sklopljeno"] = dataframed.at[i, "Naziv"] + ", " + dataframed.at[i, "Adresa"] + ", " + "OIB: " + str(dataframed.at[i, "Porezni broj"])

dataframeac = dataframee.merge(dataframef,on="Zgrada id",how="left")
dataframeaz = dataframeaz.fillna("")
dataframeac = dataframeac.merge(dataframeaz,on="Zgrada id",how="left")
dataframeac["Naziv vrste uporabe"]=""
dataframeac = dataframeac.fillna("")

for i in range(0,len(dataframeac["Cestica id"])):
    if dataframeac.at[i, "Adresa opisna"] != "":
        dataframeac.at[i, "Naziv vrste uporabe"] = dataframeac.at[i, "Naziv vrste zgrade"]+", "+dataframeac.at[i, "Adresa opisna"]
    else:
        dataframeac.at[i, "Naziv vrste uporabe"] = dataframeac.at[i, "Naziv vrste zgrade"]

dataframeac = dataframeac.merge(dataframeb,on="Cestica id",how="left")
dataframeac.drop(columns="Adresa opisna")
dataframeac["Adresa"]=""
dataframeac = dataframeac[["Posjedovni list broj","Cestica id","Naziv vrste uporabe","Povrsina","Broj cestice","Adresa"]]

dataframeab = dataframea.merge(dataframeb,on="Cestica id",how="left")
dataframeab = dataframeab.append(dataframeac,ignore_index=True)
dataframeab = dataframeab.sort_values(["Posjedovni list broj","Broj cestice"],ignore_index=True)

i = 0
j = 1
for z in range(0,len(dataframeab["Posjedovni list broj"])-1):
    if dataframeab.at[i,"Posjedovni list broj"] == dataframeab.at[i+j,"Posjedovni list broj"] and dataframeab.at[i,"Broj cestice"] == dataframeab.at[i+j,"Broj cestice"]:
        dataframeab.at[i + j, "Posjedovni list broj"] = ""
        dataframeab.at[i + j, "Adresa"] = ""
        j = j + 1
        i = i
    else:
        i = i + j
        j = 1


dataframeab = dataframeab.merge(dataframed, on="Posjedovni list broj", how = "left")
dataframeab = dataframeab.sort_values(["Broj cestice","Posjedovni list broj","Naziv vrste uporabe"],ignore_index=True,ascending=[True,False,True])
dataframeab = dataframeab[["Posjedovni list broj","Broj cestice","Naziv vrste uporabe","Omjer","Sklopljeno","Povrsina","Adresa_x"]]

i = 0
j = 1
for z in range(0,len(dataframeab["Posjedovni list broj"])-1):
    if dataframeab.at[i,"Posjedovni list broj"] == dataframeab.at[i+j,"Posjedovni list broj"] and dataframeab.at[i,"Broj cestice"] == dataframeab.at[i+j,"Broj cestice"] and dataframeab.at[i, "Naziv vrste uporabe"] == dataframeab.at[i + j, "Naziv vrste uporabe"] and dataframeab.at[i, "Povrsina"] == dataframeab.at[i + j, "Povrsina"] and dataframeab.at[i, "Adresa_x"] == dataframeab.at[i + j, "Adresa_x"]:
        dataframeab.at[i + j, "Posjedovni list broj"] = ""
        dataframeab.at[i + j, "Broj cestice"] = ""
        dataframeab.at[i + j, "Naziv vrste uporabe"] = ""
        dataframeab.at[i + j, "Povrsina"] = ""
        dataframeab.at[i + j, "Adresa_x"] = ""
        j = j + 1
        i = i
    else:
        i = i + j
        j = 1

for z in range(0,len(dataframeab["Posjedovni list broj"])-1):
    if dataframeab.at[z,"Posjedovni list broj"] == "":
        dataframeab.at[z, "Broj cestice"] = ""

for z in range(0,10):
    for i in range(0,len(dataframeab["Posjedovni list broj"])-1):
        if dataframeab.at[i,"Posjedovni list broj"] == "" and dataframeab.at[i, "Naziv vrste uporabe"] != "" and dataframeab.at[i, "Povrsina"] != "" and dataframeab.at[i-1, "Naziv vrste uporabe"] == "":
            dataframeab.at[i-1, "Naziv vrste uporabe"] = dataframeab.at[i, "Naziv vrste uporabe"]
            dataframeab.at[i-1, "Povrsina"] = dataframeab.at[i, "Povrsina"]
            dataframeab.at[i, "Povrsina"] = ""
            dataframeab.at[i, "Naziv vrste uporabe"] = ""
        else:
            pass

dataframeab = dataframeab.fillna("")
dataframeab["Drop"] = None

for i in range(0,len(dataframeab["Posjedovni list broj"])):
    if dataframeab.at[i,"Posjedovni list broj"] == "" and dataframeab.at[i, "Broj cestice"] == "" and dataframeab.at[i, "Naziv vrste uporabe"] == "" and dataframeab.at[i, "Omjer"] == "" and dataframeab.at[i, "Sklopljeno"] == "" and dataframeab.at[i, "Povrsina"] == "" and dataframeab.at[i, "Adresa_x"] == "":
        dataframeab.at[i,"Drop"] = None
    else:
        dataframeab.at[i,"Drop"] = "1"

dataframeab = dataframeab.dropna(subset=["Drop"])

#######################################################################################################################################################

dataframexzk = dataframexzk.merge(dataframeazk, on="Zk cestica id", how="left")
dataframexzk = dataframexzk[["Zk cestica id","Broj cestice brojnik","Podbroj cestice","Ulozak id","Povrsina zgrade","Redni broj ZK tijela cestice","Naziv vrste zgrade"]]
dataframeazk = dataframeazk.append(dataframexzk,ignore_index=True)
dataframeazk = dataframeazk.merge(dataframezzk, on="Zk cestica id", how="left")
dataframeabzk = dataframeazk.merge(dataframebzk, on="Ulozak id", how="left")
dataframeabzk = dataframeabzk.merge(dataframeczk, on="Zk cestica id", how="outer")
dataframedzk = dataframedzk.merge(dataframeezk, on= "Osoba upisa id", how = "left")


dataframedzk["Nazadr"] =""
dataframedzk["Omjer"] =""
dataframedzk = dataframedzk.fillna("")

for i in range(0,len(dataframedzk["Osoba upisa id"])):
    dataframedzk.at[i, "Omjer"] = str(dataframedzk.at[i, "Udio brojnik"]) + "/" + str(dataframedzk.at[i, "Udio nazivnik"])
    if dataframedzk.at[i,"Porezni broj"] != "":
        dataframedzk.at[i,"Nazadr"] = (dataframedzk.at[i,"Naziv"]+", "+ "OIB: "+dataframedzk.at[i,"Porezni broj"]+", "+str(dataframedzk.at[i,"Adresa"])).upper()
    elif dataframedzk.at[i,"Porezni broj"] == "" and dataframedzk.at[i, "Adresa"] != "":
        dataframedzk.at[i, "Nazadr"] = (dataframedzk.at[i, "Naziv"] +", "+ str(dataframedzk.at[i, "Adresa"])).upper()
    else:
        dataframedzk.at[i, "Nazadr"] = str(dataframedzk.at[i, "Naziv"]).upper()

for i in range(0,len(dataframedzk["Osoba upisa id"])):
    if  dataframedzk.at[i,"Redni broj etaze"] != "" and dataframedzk.at[i,"Udio dijela brojnik"] == "0" and dataframedzk.at[i,"Udio dijela nazivnik"] == "0":
        dataframedzk.at[i, "Naziv"] = dataframedzk.at[i, "Omjer"] +" "+ dataframedzk.at[i, "Nazadr"]
        dataframedzk.at[i,"Omjer"] = "(E-"+dataframedzk.at[i,"Redni broj etaze"]+") NEODREĐEN OMJER"
    elif dataframedzk.at[i,"Redni broj etaze"] != "" and dataframedzk.at[i,"Udio dijela brojnik"] != "":
        dataframedzk.at[i, "Naziv"] = dataframedzk.at[i, "Omjer"] +" "+ dataframedzk.at[i, "Nazadr"]
        dataframedzk.at[i, "Omjer"] = dataframedzk.at[i,"Udio dijela brojnik"]+"/"+dataframedzk.at[i,"Udio dijela nazivnik"]+" ETAŽNO VLASNIŠTVO (E-"+dataframedzk.at[i,"Redni broj etaze"]+")"
    else:
        dataframedzk.at[i,"Omjer"] = dataframedzk.at[i,"Udio dijela brojnik"]+"/"+dataframedzk.at[i,"Udio dijela nazivnik"]
        dataframedzk.at[i,"Naziv"] =  dataframedzk.at[i,"Nazadr"]

dataframedzk.to_excel(file_path+"/outputfinalnidzk.xlsx")
dataframedzk = dataframedzk.sort_values(["Redni broj ZK tijela vlasnika"],ignore_index=True)
dataframeabzk=dataframeabzk.fillna("")

for i in range(0,len(dataframeabzk["Zk cestica id"])):
    if dataframeabzk.at[i,"Podbroj cestice"] != "":
        dataframeabzk.at[i,"Broj cestice"] = dataframeabzk.at[i,"Broj cestice brojnik"]+"/"+dataframeabzk.at[i,"Podbroj cestice"]
    else:
        dataframeabzk.at[i, "Broj cestice"] = dataframeabzk.at[i, "Broj cestice brojnik"]
for i in range(0,len(dataframeabzk["Broj cestice"])):
    if dataframeabzk.at[i,"Povrsina zgrade"] != "" and dataframeabzk.at[i,"Naziv vrste zgrade"] != "":
        dataframeabzk.at[i,"Povrsina vrste uporabe"] = dataframeabzk.at[i,"Povrsina zgrade"]
        dataframeabzk.at[i,"Naziv vrste uporabe"] = dataframeabzk.at[i,"Naziv vrste zgrade"]
for i in range(0,len(dataframeabzk["Broj cestice"])):
    if dataframeabzk.at[i, "Broj cestice"].endswith("ZGR"):
        dataframeabzk.at[i, "Broj cestice"] = "*" + str(dataframeabzk.at[i, "Broj cestice"])[:-4]
    elif dataframeabzk.at[i, "Broj cestice"].endswith("zgr"):
        dataframeabzk.at[i, "Broj cestice"] = "*" + str(dataframeabzk.at[i, "Broj cestice"])[:-4]
    else:
        dataframeabzk.at[i, "Broj cestice"] = str(dataframeabzk.at[i, "Broj cestice"])

dataframeabzk = dataframeabzk.sort_values(["Broj cestice","Broj uloska","Naziv vrste uporabe"],ignore_index=True)

i = 0
j = 1
for z in range(0,len(dataframeabzk["Ulozak id"])-1):
    if dataframeabzk.at[i,"Broj cestice"] == dataframeabzk.at[i+j,"Broj cestice"] and dataframeabzk.at[i,"Broj uloska"] == dataframeabzk.at[i+j,"Broj uloska"]:
        dataframeabzk.at[i+j, "Ulozak id"] = ""
        dataframeabzk.at[i + j, "Broj cestice"] = ""
        dataframeabzk.at[i+j,"Broj uloska"] = ""
        j = j + 1
    else:
        i = i + j
        j = 1


dataframeabzk = dataframeabzk.merge(dataframedzk, on="Ulozak id", how="left")
dataframeabzk = dataframeabzk[["Broj uloska","Broj cestice","Omjer","Naziv","Naziv vrste uporabe","Povrsina vrste uporabe","Povrsina","Adresa opisna","Redni broj ZK tijela cestice","Redni broj ZK tijela vlasnika"]]

i = 0
j = 1
for z in range(0,len(dataframeabzk["Broj uloska"])-1):
    if dataframeabzk.at[i,"Broj cestice"] == dataframeabzk.at[i+j,"Broj cestice"] and dataframeabzk.at[i,"Broj uloska"] == dataframeabzk.at[i+j,"Broj uloska"]:
        dataframeabzk.at[i + j, "Broj cestice"] = ""
        dataframeabzk.at[i+j,"Broj uloska"] = ""
        dataframeabzk.at[i+j,"Povrsina"] = ""
        j = j + 1
    else:
        i = i + j
        j = 1


i = 0
j = 1
for z in range(0,len(dataframeabzk["Broj uloska"])-1):
    if dataframeabzk.at[i,"Naziv vrste uporabe"] == dataframeabzk.at[i+j,"Naziv vrste uporabe"] and dataframeabzk.at[i,"Povrsina vrste uporabe"] == dataframeabzk.at[i+j,"Povrsina vrste uporabe"]:
        dataframeabzk.at[i+j, "Naziv vrste uporabe"] = ""
        dataframeabzk.at[i+j,"Povrsina vrste uporabe"] = ""
        j = j + 1
    else:
        i = i + j
        j = 1

i = 0
j = 1
for z in range(0,len(dataframeabzk["Broj uloska"])-1):
    if dataframeabzk.at[i,"Adresa opisna"] == dataframeabzk.at[i+j,"Adresa opisna"] and dataframeabzk.at[i+j,"Broj cestice"] == "":
        dataframeabzk.at[i+j, "Adresa opisna"] = ""
        j = j + 1
    else:
        i = i + j
        j = 1

dataframeabzk = dataframeabzk.fillna("")
for i in range(0,len(dataframeabzk["Broj uloska"])-1):
    if dataframeabzk.at[i,"Naziv vrste uporabe"] != "" and dataframeabzk.at[i,"Naziv"] == "":
        dataframeabzk.at[i, "Povrsina"] = ""
        i = i + 1
    else:
        pass

for i in range(0,len(dataframeabzk["Broj uloska"])-1):
    if dataframeabzk.at[i,"Broj cestice"] == "":
        dataframeabzk.at[i, "Redni broj ZK tijela cestice"] = ""
        i = i + 1
    else:
        pass

for z in range(0,50):
    for i in range(0,len(dataframeabzk["Broj uloska"])-1):
        if dataframeabzk.at[i,"Broj uloska"] == "" and dataframeabzk.at[i, "Naziv vrste uporabe"] != "" and dataframeabzk.at[i, "Povrsina vrste uporabe"] != "" and dataframeabzk.at[i-1, "Naziv vrste uporabe"] == "":
            dataframeabzk.at[i-1, "Naziv vrste uporabe"] = dataframeabzk.at[i, "Naziv vrste uporabe"]
            dataframeabzk.at[i-1, "Povrsina vrste uporabe"] = dataframeabzk.at[i, "Povrsina vrste uporabe"]
            dataframeabzk.at[i, "Povrsina vrste uporabe"] = ""
            dataframeabzk.at[i, "Naziv vrste uporabe"] = ""
        else:
            pass

dataframeabzk["Drop"] = None
for i in range(0,len(dataframeabzk["Broj uloska"])):
    if dataframeabzk.at[i,"Broj uloska"] == "" and dataframeabzk.at[i, "Broj cestice"] == "" and dataframeabzk.at[i, "Omjer"] == "" and dataframeabzk.at[i, "Naziv"] == "" and dataframeabzk.at[i, "Naziv vrste uporabe"] == "" and dataframeabzk.at[i, "Povrsina vrste uporabe"] == "" and dataframeabzk.at[i, "Povrsina"] == "":
        dataframeabzk.at[i,"Drop"] = None
    else:
        dataframeabzk.at[i,"Drop"] = "1"
dataframeabzk = dataframeabzk.dropna(subset=["Drop"])
cestcut = pd.DataFrame()
cestcut = dataframeab[["Broj cestice"]].copy()
cestcut.to_excel(file_path+"/cestcut.xlsx")
cestzk = []
j = 1


################################## Ako treba nešto izbaciti
#dataframeabzk = dataframeabzk.reset_index(drop = True)
#dataframeabzk = dataframeabzk.drop([58,59,60,61,62])
##################################


dataframeabzk = dataframeabzk.reset_index(drop=True)
for i in range(0,len(dataframeabzk["Broj cestice"])):
    if dataframeabzk.at[i,"Broj cestice"] != "":
        dataframeabzk.at[i,"Drop"] = j
        j = j+1
        cestzk.append(dataframeabzk.at[i,"Broj cestice"])
    else:
        dataframeabzk.at[i,"Drop"] = ""

for i in range(0,len(dataframeabzk["Broj cestice"])):
    if dataframeabzk.at[i,"Naziv vrste uporabe"] == "":
        dataframeabzk.at[i,"Naziv vrste uporabe"] = dataframeabzk.at[i,"Adresa opisna"]
    else:
        pass
del dataframeabzk["Adresa opisna"]

def duplicipos(i):
    if dataframeab.at[i,"Broj cestice"] in lista:
        return True
    else:
        return False
def cestzkfun(i):
    if dataframeab.at[i,"Broj cestice"] in cestzk:
        return True
    else:
        return False

dataframeab = dataframeab.reset_index(drop = True)
lista=[]

j = 1
for i in range(0,len(dataframeab["Broj cestice"])):
    if dataframeab.at[i,"Broj cestice"] != "" and duplicipos(i) == False and cestzkfun(i) == True:
        dataframeab.at[i,"Drop"] = j
        j = j+1
        lista.append(dataframeab.at[i,"Broj cestice"])
    else:
        dataframeab.at[i,"Drop"] = ""


dataframeabzk.to_excel(file_path+"/outputfinalniabzk.xlsx")
dataframeab.to_excel(file_path+"/outputfinalniab.xlsx")
wbab = load_workbook(file_path+"/outputfinalniab.xlsx")
wbabzk = load_workbook(file_path+"/outputfinalniabzk.xlsx")
wsbab = wbab.active
wsbabzk = wbabzk.active

if len(dataframeab["Broj cestice"]) > len(dataframeabzk["Broj cestice"]):
    a = len(dataframeab["Broj cestice"])
else:
    a = len(dataframeabzk["Broj cestice"])


for i in range(2,a+100):
    if wsbab.cell(row=i, column=9).value != wsbabzk.cell(row=i, column=11).value and wsbab.cell(row=i, column=9).value == None and wsbabzk.cell(row=i, column=11).value != None:
        wsbabzk.insert_rows(i)
        i = i
    if wsbab.cell(row=i, column=9).value != wsbabzk.cell(row=i, column=11).value and wsbab.cell(row=i, column=9).value != None and wsbabzk.cell(row=i, column=11).value == None:
        wsbab.insert_rows(i)
        i = i
    else:
        i = i + 1

masterpos = pd.DataFrame(wsbab.values)
masterzk = pd.DataFrame(wsbabzk.values)
masterpos.to_excel(file_path+"/masterpos.xlsx")
masterzk.to_excel(file_path+"/masterzk.xlsx", index=False)
masterpos.columns = masterpos.iloc[0]

masterpos = masterpos.drop(masterpos.index[0])
masterpos = masterpos.set_index(["Drop"])
masterpos.columns = ["A","Broj posjedovnog lista","Broj cestice","Način uporabe katastarske čestice","Omjer","Prezime i ime odnosno tvrtka ili naziv upisane osobe.Prebivalište odnosno sjedište, ulica , kućni broj i OIB upisane osobe (posjednika).","Povrsina [m2]","Adresa"]
del masterpos["A"]

masterzk.columns = masterzk.iloc[0]
masterzk = masterzk.drop(masterzk.index[0])
masterzk.columns = ["A","Broj Z.K. uloška","Broj katastarske čestice","Omjer","Prezime i ime odnosno tvrtka ili naziv upisane osobe.Prebivalište odnosno sjedište, ulica , kućni broj i OIB upisane osobe (vlasnika).","Način uporabe katastarske čestice","Površina vrste uporabe", "Površina ukupna [m2]","Redni broj ZK tijela cestice","Redni broj ZK tijela vlasnika","Drop"]
del masterzk["A"]
del masterzk["Drop"]


masterpos.to_excel(file_path+"/masterpos.xlsx")
masterzk.to_excel(file_path+"/masterzk.xlsx", index=False)
ws = load_workbook(file_path+"/masterpos.xlsx")
wss = load_workbook(file_path+"/masterzk.xlsx")
wsa = ws.active
wsb = wss.active
wsa.column_dimensions['A'].width = 9
wsa.column_dimensions['B'].width = 13
wsa.column_dimensions['C'].width = 13
wsa.column_dimensions['D'].width = 13
wsa.column_dimensions['E'].width = 10
wsa.column_dimensions['F'].width = 50
wsa.column_dimensions['G'].width = 13
wsa.column_dimensions['H'].width = 13

for row in wsa.iter_rows():
    for cell in row:
        cell.alignment = cell.alignment.copy(wrapText=True, horizontal="center", vertical="center")

wsb.column_dimensions['A'].width = 13
wsb.column_dimensions['B'].width = 13
wsb.column_dimensions['C'].width = 10
wsb.column_dimensions['D'].width = 50
wsb.column_dimensions['E'].width = 13
wsb.column_dimensions['F'].width = 13
wsb.column_dimensions['G'].width = 13
for row in wsb.iter_rows():
    for cell in row:
        cell.alignment = cell.alignment.copy(wrapText=True, horizontal="center", vertical="center")
ws.save(file_path+"/masterpos.xlsx")
wss.save(file_path+"/masterzk.xlsx")













