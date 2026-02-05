# Import libraries
import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO

output = BytesIO()
today = date.today().strftime("%d%m%Y")

# Insert title
st.title("Conv Colisage Mandiang")

# Upload the file
uploaded_file = st.file_uploader("Carica un nuovo colisage", type = ["xls", "xlsx"])

# If the file exists
if uploaded_file is not None:
    # Import the dataframe
    df = pd.read_excel(uploaded_file, skiprows = 4)
    df = df.rename(columns={"DESIGNATION": "pesce",
                            "Nombre de Coli": "numero_casse",
                            "Poids net par Coli ": "peso_per_cassa",
                            "DETAIL": "dettaglio"})
    df = df[["pesce", "numero_casse", "peso_per_cassa", "dettaglio"]]
    df = df.dropna(axis = 0, how = 'any', subset = None, inplace = False)
    L_numero = []
    L_pesce = []
    L_info = []
    L_peso = []
    
    
    def split_digit (text):
        splitted = []  
        tmp = []       

        for c in text:
            if c.isdigit():
                tmp.append(c)   

            elif tmp:           
                splitted.append(''.join(tmp))
                tmp = []

        if tmp:
            splitted.append(''.join(tmp))
        return splitted


    index = 1 
    for pesce, num, peso, dett in zip(df.pesce, df.numero_casse, df.peso_per_cassa, df.dettaglio):
        num = int(num)
        num_list = split_digit(dett.split("(")[0])
        
        if num == 1:
            L_numero.append(int(num_list[0]))
            L_pesce.append(pesce)
            L_peso.append(peso)
            
            try:
                val = dett.split('(', 1)[1].split(')')[0]
            except IndexError:
                val = ""
                
            L_info.append(val)
            
        else:
            if len(num_list) == num:
                L_numero.extend(int(x) for x in num_list )  
            else:
                L_numero.extend(list(range(int(num_list[0]), int(num_list[-1])+1)))
                
            L_pesce.extend([pesce]*num)
            L_peso.extend([peso]*num)
            
            try:
                val = dett.split('(', 1)[1].split(')')[0]
            except IndexError:
                val = ""
                
            L_info.extend([val]*num)


    new_df = pd.DataFrame(None)
    new_df["Numero"] = L_numero
    new_df["Pesce"] = L_pesce
    new_df["Pezzi"] = L_info
    new_df["Cliente"] = [""]*len(L_numero)
    new_df["Peso"] = L_peso
    new_df = new_df.sort_values(by = ["Numero"])
    new_df = new_df.set_index("Numero")


    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    # Write each dataframe to a different worksheet.
        new_df.to_excel(writer)
        writer.close()


    st.download_button(
       label = "Clicca per Scaricare",
       data = output,
       file_name = "colisage_" + str(today) + ".xlsx",
       mime = "application/vnd.ms-excel"
    )
    

    


