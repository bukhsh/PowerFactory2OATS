#__author__ = "W. Bukhsh"
# This script is executed within DIgSILENT PowerFactory, and outputs a bunch of csvfiles


#read csv data into pandas dataframes
import pandas as pd
terminals            = pd.DataFrame.from_csv('data/csv/terminals.csv', index_col=None)
loads                = pd.DataFrame.from_csv('data/csv/loads.csv', index_col=None)
generators           = pd.DataFrame.from_csv('data/csv/generators.csv', index_col=None)
branches             = pd.DataFrame.from_csv('data/csv/branches.csv', index_col=None)
transformers2w       = pd.DataFrame.from_csv('data/csv/transformers.csv', index_col=None)
transformers2w_extra = pd.DataFrame.from_csv('data/csv/transformers_char.csv', index_col=None)
transformers3w       = pd.DataFrame.from_csv('data/csv/transformers3.csv', index_col=None)
transformers3w_extra = pd.DataFrame.from_csv('data/csv/transformers3_char.csv', index_col=None)
busbarswitches       = pd.DataFrame.from_csv('data/csv/busbarswitches.csv', index_col=None)

#Process data

#=======TERMINAL DATA=======
terminals_df = pd.DataFrame(columns={'bus_ID','bus_name','basekV'})
ind = 0
for i in terminals.index.tolist():
    terminals_df.loc[ind] = pd.Series({'bus_name':terminals['terminal'][i],'basekV':terminals['basekV'][i]})
    ind += 1
#print len(bus)
terminals_df = terminals_df.drop_duplicates(['bus_name'])

#=======SWITCHES=======
ind =0
switches = pd.DataFrame(columns={'name','bus_from','bus_to','status'})
for i in busbarswitches.index.tolist():
    switches.loc[ind] = pd.Series({'name':busbarswitches['name'][i],'bus_from':busbarswitches['bus_from'][i],
    'bus_to':busbarswitches['bus_to'][i],'status':busbarswitches['status'][i]})
    ind += 1
#delete switches that are switched off, and also that does not include 'to' or 'from' bus
switches = switches[switches['status']!=0]
switches = switches[switches['bus_from']!='None']
switches = switches[switches['bus_to']!='None']
#reset the index
switches = switches.reset_index(drop=True)
count = 0
for i in switches.index.tolist():
    if switches['bus_from'][i] not in terminals_df['bus_name'] or switches['bus_to'][i] not in terminals_df['bus_name']:
        count += 1

#=======branch data=======
branch_df = pd.DataFrame(columns={'branch_ID','from_bus_ID','to_bus_ID','r','x','rate'})
for i in branches.index.tolist():
    branch_df.loc[ind] = pd.Series({'branch_ID':branches['branch_ID'][i],\
    'from_bus_ID':branches['from_bus_ID'][i],'to_bus_ID':branches['to_bus_ID'][i],\
    'r':branches['r'][i],'x':branches['x'][i],'rate':branches['rate'][i]})
    ind += 1

#=======2-winding transformer data=======
transformer2w_df = pd.DataFrame(columns={'transformer_ID','from_bus_ID','to_bus_ID','r','x','base','rate'})
for i in transformers2w.index.tolist():
    #print transformers2w['transformer_ID'][i]
    #print 'match: ',transformers2w_extra['r'][transformers2w_extra['transformer_ID']== transformers2w['transformer_ID'][i]].values
    r    = transformers2w_extra['r'][transformers2w_extra['transformer_ID']==transformers2w['transformer_ID'][i]].values
    x    = transformers2w_extra['x'][transformers2w_extra['transformer_ID']==transformers2w['transformer_ID'][i]].values
    base = transformers2w_extra['base '][transformers2w_extra['transformer_ID']==transformers2w['transformer_ID'][i]].values
    transformer2w_df.loc[ind] = pd.Series({'transformer_ID':transformers2w['transformer_ID'][i],\
    'from_bus_ID':transformers2w['from_bus_ID'][i],'to_bus_ID':transformers2w['to_bus_ID'][i],\
    'r':r,'x':x,'base':base,'rate':transformers2w['rate'][i]})
    ind += 1

#=======3-winding transformer data=======
transformer3w_df = pd.DataFrame(columns={'transformer_ID','HV_bus_ID','MV_bus_ID',\
'LV_bus_ID','rHV','xHV','rMV','xMV','rLV','xLV','baseHV','baseMV','baseLV','rateHV','rateMV','rateLV'})
for i in transformers3w.index.tolist():
    rHV    = transformers3w_extra['rHV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values
    xHV    = transformers3w_extra['xHV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values
    rMV    = transformers3w_extra['rMV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values
    xMV    = transformers3w_extra['xMV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values
    rLV    = transformers3w_extra['rLV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values
    xLV    = transformers3w_extra['xLV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values
    baseHV = transformers3w_extra['baseHV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values
    baseMV = transformers3w_extra['baseMV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values
    baseLV = transformers3w_extra['baseLV'][transformers3w_extra['transformer_ID']==transformers3w['transformer_ID'][i]].values

    transformer3w_df.loc[ind] = pd.Series({'transformer_ID':transformers3w['transformer_ID'][i],\
    'HV_bus_ID':transformers3w['HV_bus_ID'][i],'MV_bus_ID':transformers3w['MV_bus_ID'][i],\
    'LV_bus_ID':transformers3w['LV_bus_ID'][i],'rHV':rHV,'xHV':xHV,'rMV':rMV,'xMV':xMV,\
    'rLV':rLV,'xLV':xLV,'baseHV':baseHV,'baseMV':baseMV,'baseLV':baseLV,'rateHV':transformers3w['rateHV'][i],\
    'rateMV':transformers3w['rateMV'][i],'rateLV':transformers3w['rateLV'][i]})
    ind += 1


#=======Write data into a spreedsheet=======
writer = pd.ExcelWriter('data.xlsx')
switches.to_excel(writer,'switches',index=False)
terminals_df.to_excel(writer,'terminals',index=False)
generators.to_excel(writer,'generators',index=False)
loads.to_excel(writer,'load',index=False)
branch_df.to_excel(writer,'branch',index=False)
transformer2w_df.to_excel(writer,'two-winding transformers',index=False)
transformer3w_df.to_excel(writer,'three-winding transformers',index=False)
writer.save()
