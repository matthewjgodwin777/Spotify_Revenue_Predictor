import numpy as np
import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from datetime import datetime
from sklearn.model_selection import train_test_split

path=r'C:\Users\Administrator\Desktop\RNspotifydataset.xlsx'
path2=r'C:\Users\Administrator\Desktop\RNspotifytestset.xlsx'
X=[]
Y=[]
country=[]
refval=[]

def setval(ct,er,st):
    if ct in country:
        pos=country.index(ct)
        eps1,st1=tuple(refval[pos])
        eps=er/st
        neweps=round(((eps*st)+(eps1*st1))/(st+st1),6)
        del country[pos]
        del refval[pos]
        pos=0
        for i in range(len(country)):
            if neweps<=refval[i][0]:
                pos=i
                break
        country.insert(pos,ct)
        refval.insert(pos,[neweps,int(st+st1)])
    else:
        pos=0
        eps=round(er/st,6)
        for i in range(len(country)):
            if eps<=refval[i][0]:
                pos=i
                break
        country.insert(pos,ct)
        refval.insert(pos,[eps,int(st)])
#Extracting First Sheet of Excel File alone
sheet_count = len(pd.ExcelFile(path).sheet_names)
doc = pd.read_excel(path,sheet_name=list(range(sheet_count)))
doc=doc[0]
rows=len(doc['Stream'])
for i in range(rows):
    if doc['Stream'][i]!=0:
        setval(doc['Customer Territory'][i],doc['Earnings($)'][i],doc['Stream'][i])
for i in range(rows):
    if doc['Stream'][i]!=0:
            #X.append([refval[country.index(doc['Customer Territory'][i])][0],doc['Stream'][i]])
            X.append([doc['Earnings($)'][i]/doc['Stream'][i],doc['Stream'][i]]) #replace above line and comment this line for 2nd version of ML model outputs
            Y.append(doc['Earnings($)'][i]/doc['Stream'][i])


#x_train, x_test,y_train,y_test = train_test_split(X,Y,test_size =0.05)
x_train=X
y_train=Y
x_test=[]

y_train = np.array(y_train)

rf = RandomForestRegressor(n_estimators=100, random_state=int(datetime.now().strftime("%S")))

x_train=np.array(x_train)
rf.fit(x_train, y_train)

####
sheet_count = len(pd.ExcelFile(path2).sheet_names)
doc = pd.read_excel(path2,sheet_name=list(range(sheet_count)))
doc=doc[0]
rows=len(doc['Stream'])
cnames=[]
streams=[]
notacc=[]
REALCOST=0.0
tm=0
for i in range(rows):
    if doc['Customer Territory'][i] not in country:
        notacc.append(i)
    elif 'Spotify' in doc['Retailer'][i]:
        x_test.append([refval[country.index(doc['Customer Territory'][i])][0],int(doc['Stream'][i])])
        cnames.append(doc['Customer Territory'][i])
        streams.append(doc['Stream'][i])
        REALCOST+=doc['Earnings($)'][i]
    elif not pd.isna(doc['Earnings($)'][i]):
        tm+=1
y_pred = rf.predict(np.array(x_test))
print('Predicted Estimates: -\n')
print('Country','\t','Streams','\t','Estimate')
print('------------------------------------')
est=0.0
for i in range(len(y_pred)):
    print(cnames[i],'\t',int(streams[i]),'\t',y_pred[i]*x_test[i][1])
    est+=y_pred[i]*x_test[i][1]
print('--------------------------------------------\n')
print('\nUnaccounted Entries: -')
tm=0
for x in notacc:
    try:
        print(doc['Customer Territory'][x],'\t',int(doc['Stream'][x]))
        tm+=1
    except:
        pass
if tm==0:
    print("No entries\n")
print('--------------------------------------------\nPredicted Estimate: ',est)
print('Actual Cost ($) : ',REALCOST)
####

