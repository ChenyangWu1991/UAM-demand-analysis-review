import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
plt.rcParams.update({'font.size': 10, 'font.family': 'Arial'})
Odd=pd.read_excel('Odd ratio summary.xlsx',usecols=[0,1,2,3,4])
# Set the start point of the four categories
Odd['Length']=Odd.iloc[:,3]+Odd.iloc[:,4]
Odd['Mixed_start']=Odd['Length'].apply(lambda x:-round(x/2))
Odd['Not_sig_start']=Odd['Mixed_start']+Odd['Mixed influence number']
Odd['Negative_start']=Odd['Mixed_start']-Odd.iloc[:,2]
Odd['Positive_start']=Odd['Mixed_start']+Odd['Length']
print(Odd)
# Print the horizontal bar fig
fig=plt.figure(figsize=(10,6.5))
plt.barh(Odd['Attribute'],left=Odd['Negative_start'],width=Odd.iloc[:,2],
         height=0.5,color='#ff9999', edgecolor='black',label='Negative')
plt.barh(Odd['Attribute'],left=Odd['Mixed_start'],width=Odd.iloc[:,4],
         height=0.5,color='#66b3ff', edgecolor='black',label='Mixed')
plt.barh(Odd['Attribute'],left=Odd['Not_sig_start'],width=Odd.iloc[:,3],
         height=0.5,color='gold', edgecolor='black',label='Not Significant')
plt.barh(Odd['Attribute'],left=Odd['Positive_start'],width=Odd.iloc[:,1],
         height=0.5,color='#99ff99', edgecolor='black',label='Positive')
plt.xlim(-55,45)
plt.axvline(x=0, linestyle='--',color='black')
plt.xticks([i for i in range(-55,51,5)],[55,50,45,40,35,30,25,20,15,10,5,0,5,10,15,20,25,30,35,40,45,50])
plt.xlabel('Number of papers',fontsize=12)
plt.ylabel('Factors',fontsize=12,labelpad=40)
plt.legend()
plt.tight_layout()
plt.gca().invert_yaxis()
plt.savefig('Odd.svg', dpi=800)
plt.show()
