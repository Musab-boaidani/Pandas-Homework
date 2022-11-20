import pandas as pd

files=[]
filesFolser='Excel files'
finalDF=pd.DataFrame()


for i in range(10):
    df=pd.read_excel(f"{filesFolser}\\{i+1}-11-2022.xlsx")
    files.append({1:df,2:f"{i+1}-11-2022"})



allBasicInfo=[]

for i in files:
    allBasicInfo.append(i[1][['Name','Email','Guest']] )

basicInfo=pd.concat(allBasicInfo)
basicInfo.drop_duplicates(subset="Name",inplace=True)

 
for n,e,g in zip(basicInfo["Name"],basicInfo["Email"],basicInfo["Guest"]) :
    sum=0
    absent=0
    dates={}
    for f in files:
        x=f[1][f[1]["Name"] == n]
        x=x.squeeze()
        if(x.empty==False):
            sum+=x["Time"]
            dates[f[2]]=x["Time"]

        else:
            absent+=1
            dates[f[2]]=0

        
    rowData = {'Name': n, 'Email': e, 'Guest': g, **dates ,"Time":sum,"Absent":absent}
    row = pd.Series(data=rowData)
    row=row.to_frame().T
    finalDF= pd.concat([finalDF,row],axis=0)

finalDF.to_excel("final.xlsx",index=False)
    

   


