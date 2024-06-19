import numpy as np
import pandas as pd
from datetime import datetime, timedelta

local='eldorado'

#dmi=pd.read_csv('dmi_'+local+'.csv',header=None,sep=";")

dmi=pd.read_excel('dmi_'+local+'.xlsx',sheet_name='Planilha1',header=None)
dmi=np.array(dmi.iloc[:,:])
vlbr=pd.read_csv('voltbras_'+local+'.csv',header=None,sep=";")
vlbr=np.array(vlbr.iloc[:,:])

#Organiza arquivo voltbras
vlbr1=np.array(vlbr[0,:]).reshape(1,-1)
for i in range(vlbr.shape[0]-1,0,-1):
    vlbr1=np.append(vlbr1,np.array(vlbr[i,:]).reshape(1,-1),axis=0)
    
vlbr=vlbr1
del vlbr1,i
vlbr=np.delete(vlbr,np.array([5,6,7,8,11,12,13,14,15,16,17,18,19]),axis=1)

############ ORGANIZA DADOS VOLTBRAS ##################
vlbr=np.append(vlbr,np.zeros([vlbr.shape[0],3]),axis=1)
vlbr[0,vlbr.shape[1]-2]="INICIO"
vlbr[0,vlbr.shape[1]-1]="FIM"
vlbr[0,vlbr.shape[1]-3]="Alocado"

for i in range(1,vlbr.shape[0]):
    p1=datetime.strptime(vlbr[i,1]+" "+vlbr[i,2], "%d/%m/%Y %H:%M")
    p2=datetime.strptime(vlbr[i,1]+" "+vlbr[i,3], "%d/%m/%Y %H:%M")
    
    if p1>p2:
        p2=p2+timedelta(days=1)
    
    fuso=3
    vlbr[i,vlbr.shape[1]-2]=p1-timedelta(hours=fuso)
    vlbr[i,vlbr.shape[1]-1]=p2-timedelta(hours=fuso)

del i,p2,p1,fuso
    
####### DICIONARIO DE DADOS MEDIDOR ###########
dicrec={}
#Laço para filtrar apenas os eventos de recarga
cont=0
i=1

while i<dmi.shape[0]:
    if float(dmi[i,2])>1000 or float(dmi[i,3])>1000 or float(dmi[i,4])>1000:
        recfilt=np.array(dmi[0,:]).reshape(1,-1)
        cont+=1
        while float(dmi[i,2])>1000 or float(dmi[i,3])>1000 or float(dmi[i,4])>1000:
            recfilt=np.append(recfilt,np.array(dmi[i,:]).reshape(1,-1),axis=0)
            i+=1
            if i==dmi.shape[0]:
                break
        recfilt=np.append(recfilt,np.array(dmi[i,:]).reshape(1,-1),axis=0)
        #dicrec[cont]={"dmi":recfilt, "dmi_energia":round(np.sum(np.array(recfilt[1:recfilt.shape[0]-1,1],dtype=float))/(120*1000),2),
         #             "dmi_horainicial":datetime.strptime(recfilt[1,0],"%d/%m/%Y %H:%M:%S"), 
          #            "dmi_horafinal":datetime.strptime(recfilt[recfilt.shape[0]-1,0],"%d/%m/%Y %H:%M:%S")}
        dicrec[cont]={"dmi":recfilt, "dmi_energia":round(np.sum(np.array(recfilt[1:recfilt.shape[0]-1,1],dtype=float))/(120*1000),2),
                      "dmi_horainicial":recfilt[1,0], 
                      "dmi_horafinal":recfilt[recfilt.shape[0]-1,0]}
    else:
        i+=1

del i,recfilt,cont

############################################################################################################################################################################
limite=100
limite2=60
for chave in dicrec:
    horai=dicrec[chave]["dmi_horainicial"]
    horaf=dicrec[chave]["dmi_horafinal"]
    org=0
    for i in range(1,vlbr.shape[0]):
        horavlbr=vlbr[i,vlbr.shape[1]-2]
        horavlbrf=vlbr[i,vlbr.shape[1]-1]
        if int(abs(horavlbr.timestamp()-horai.timestamp()))<=limite or (horavlbr.timestamp()>=horai.timestamp() and horavlbr.timestamp()<horaf.timestamp() and 
                                                                        (horavlbrf.timestamp()<=horaf.timestamp() or int(abs(horaf.timestamp()-horavlbrf.timestamp()))<=limite2)):
            vlbr[i,vlbr.shape[1]-3]=chave
            if org==0:
                dicrec[chave]['vlbr']=np.array(vlbr[i,:]).reshape(1,-1)
            elif org>0:
                dicrec[chave]['vlbr']=np.append(dicrec[chave]['vlbr'],np.array(vlbr[i,:]).reshape(1,-1),axis=0)
            org+=1
    dicrec[chave]['vlbr_eventos']=org
    if org>0:
        dicrec[chave]['vlbr_energia']=round(np.sum(np.array(dicrec[chave]['vlbr'][:,4],dtype=float)),2)
        
        
###########################################################################################################################################################################
# RELATORIO DE SAIDA

exp1=np.array(["Janela","Energia Consumida","Hora Inicial", "Hora Final","Flag Potência(0-normal,1-VB>5kWh,2-DMI>5kWh)","Consumo Voltbras"]).reshape(1,-1)
exp2=np.array(["Codigo","Energia Consumida","Hora Inicial", "Hora Final","Conector"]).reshape(1,-1)

for chave in dicrec:
    aux=int(dicrec[chave]["vlbr_eventos"])
    if aux==0:
        exp1=np.append(exp1,np.array([chave,dicrec[chave]["dmi_energia"],dicrec[chave]["dmi_horainicial"],dicrec[chave]["dmi_horafinal"],0,0.0]).reshape(1,-1),axis=0)
    elif aux>0 and (dicrec[chave]["vlbr_energia"]-dicrec[chave]["dmi_energia"])>5:
        exp1=np.append(exp1,np.array([chave,dicrec[chave]["dmi_energia"],dicrec[chave]["dmi_horainicial"],dicrec[chave]["dmi_horafinal"],1,dicrec[chave]["vlbr_energia"]]).reshape(1,-1),axis=0)
    elif aux>0 and (dicrec[chave]["dmi_energia"]-dicrec[chave]["vlbr_energia"])>5:
        exp1=np.append(exp1,np.array([chave,dicrec[chave]["dmi_energia"],dicrec[chave]["dmi_horainicial"],dicrec[chave]["dmi_horafinal"],2,dicrec[chave]["vlbr_energia"]]).reshape(1,-1),axis=0)

for i in range(1,vlbr.shape[0]):
    if int(vlbr[i,vlbr.shape[1]-3])==0:
        exp2=np.append(exp2,np.array([vlbr[i,5],vlbr[i,4],vlbr[i,8],vlbr[i,9],vlbr[i,6]]).reshape(1,-1),axis=0)
        

# Definindo o caminho do arquivo de saída
file_path = 'relatorio_saida_'+local+'.xlsx'

# Usando ExcelWriter para criar um arquivo com múltiplas planilhas
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    pd.DataFrame(exp1).to_excel(writer, sheet_name='Erros_Janelas_DMI', index=False, header=False)
    pd.DataFrame(exp2).to_excel(writer, sheet_name='Erros_Recargas_Volbras', index=False, header=False)
    
