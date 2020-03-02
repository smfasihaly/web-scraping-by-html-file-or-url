from bs4 import BeautifulSoup
import requests
import pandas as pd 

from openpyxl import load_workbook
url="a.html"
lstfName = []
lstlName = []
lstphone = []
lstemail = []
lststate = []
lstcity  = []
soup = BeautifulSoup(open("a.html"), "lxml")
mytxt =soup.text
soup = BeautifulSoup(mytxt, 'lxml')
# Make a G.ET request to fetch the raw HTML content
gdp_table = soup.find("div", attrs={"member-name"})
try:
    for i in range(100):
        names = soup.find("div", attrs={'id':"MainCopy_ctl16_Contacts_DisplayNamePanel_"+str(i)})
        addresses = soup.find("div", attrs={'id':"MainCopy_ctl16_Contacts_Addr1Panel4_"+str(i)})
        country = soup.find("div", attrs={'id':"MainCopy_ctl16_Contacts_Addr1Panel5_"+str(i)})
        emails = soup.find("div", attrs={'id':"MainCopy_ctl16_Contacts_EmailAddressPanel_"+str(i)})
        phones = soup.find("div", attrs={'id':"MainCopy_ctl16_Contacts_PhonePanel_"+str(i)})
        lst = {'usa','united states', 'unitedstates'}
        if i == 63:
            print('')
        if True: #country.text.strip().lower() in lst:
            name = names.text.strip()
            try:
                city,state = addresses.text.split(',')
            except:
                city = ''
                if addresses != None:
                    city =addresses.text
                state = ''
            city = city.strip()
            state =  state.strip()
            state = state.split(' ')[0]
            
            email = '' 
            if emails != None:
                email = emails.text.strip()
        
            phone = ''
            if phones != None:
                try:
                    phone = phones.text.strip().split(' ')[0]
                except:
                    phone = phones.text.strip()
            lstfName.append(name.split(' ')[0])
            lstlName.append( "-".join(name.split(' ')[1:]))
            lststate.append(state)
            lstemail.append(email)
            lstphone.append(phone)
            lstcity.append(city)

            print( "-".join(name.split(' ')[1:]),city,state,email,phone)
    
except:
    print()
my_dict = { 'First Name' : lstfName,
                    'Last Name' : lstlName,
                    'City' : lstcity,
                    'State' : lststate,
                    'Email' : lstemail,
                    'Phone': lstphone}
df = pd.DataFrame(my_dict)
df.to_csv('file.csv')
reader =[]
writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')
# try to open an existing workbook
writer.book = load_workbook('output.xlsx')
# copy existing sheets
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
# read existing file
reader = pd.read_excel(r'output.xlsx')
# write out the new sheet
df.to_excel(writer,index=False,header=False,startrow=len(reader)+1)

writer.close()

# Parse the html content
#print(soup.text) # print the parsed data of html
