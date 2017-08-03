#__author__ = "W. Bukhsh"
# This script is executed within DIgSILENT PowerFactory, and outputs a bunch of csvfiles

#importing of pf module
import powerfactory


# Calling app Appication object
app=powerfactory.GetApplication()

#Calling ldf Command object (ComLdf)
ldf=app.GetFromStudyCase('ComLdf')

#executing the load flow command
ldf.Execute()

#======Extract information about the system======

# List of lines contained in the project
Lines        = app.GetCalcRelevantObjects('*.ElmLne')
Branches        = app.GetCalcRelevantObjects('*.ElmBranch')

#two winding transformers
Transformers2  = app.GetCalcRelevantObjects('*.ElmTr2')
Transformers2_char = app.GetCalcRelevantObjects('*.TypTr2') #need to extract resistance and reactance

#three winding transformers
Transformers3  = app.GetCalcRelevantObjects('*.ElmTr3')
Transformers3_char = app.GetCalcRelevantObjects('*.TypTr3') #need to extract resistance and reactance

#other components
Terminals    = app.GetCalcRelevantObjects('*.ElmTerm')
Cubicles     = app.GetCalcRelevantObjects('*.StaCubic')
SynGenerators   = app.GetCalcRelevantObjects('*.ElmSym')
StatGenerators  = app.GetCalcRelevantObjects('*.ElmGenStat')
Loads        = app.GetCalcRelevantObjects('*.ElmLod')
BusbarSwitches   = app.GetCalcRelevantObjects('*.ElmCoup')
Substations = app.GetCalcRelevantObjects('*.ElmSubstat')


#Write data into csvfiles

#Branch data
flines = open('data/csv/branches.csv','w')
flines.write('branch_ID,from_bus_ID,to_bus_ID,r,x,rate\n')
#for branch in Branches: #get each element out of list
for branch in Lines: #get each element out of list
    name=branch.loc_name # get name of the line
    bus1 = branch.bus1.loc_name#cTerm0 # get name of the line
    bus2 = branch.bus2.loc_name#cTerm1 # get name of the line
    try:
        r   = branch.R1
    except:
        r = 'None'
    try:
        x   = branch.X1
    except:
        x = 'None'
    try:
        rate = branch.ContRating
    except:
        rate = 'None'
    flines.write('%s, %s, %s, %s, %s, %s \n' %(name,bus1,bus2,r,x,rate))
    #print to file
flines.close()


#=====Two winding transformers=====
#transformer data
ftrans = open('data/csv/transformers.csv','w')
ftrans.write('transformer_ID,from_bus_ID,to_bus_ID,rate\n')
for trans in Transformers2: #get each element out of list
    name=trans.loc_name # get name of the line
    bus1=trans.bushv.cterm.loc_name # get name of the line
    bus2=trans.buslv.cterm.loc_name # get name of the line
    rate = trans.Snom_a
    ftrans.write('%s, %s, %s, %s \n' %(name,bus1,bus2,rate))
ftrans.close()

#transformer data for extracting resistance and reactance
ftrans2 = open('data/csv/transformers_char.csv','w')
ftrans2.write('transformer_ID,r,x,base \n')
for trans in Transformers2_char: #get each element out of list
    name=trans.loc_name # get name of the line
    r   = trans.r1pu
    x   = trans.x1pu
    base = trans.strn
    ftrans2.write('%s, %s, %s, %s \n' %(name,r,x, base))
ftrans2.close()

#=====Three winding transformers=====
#transformer data
ftrans = open('data/csv/transformers3.csv','w')
ftrans.write('transformer_ID,HV_bus_ID,MV_bus_ID,LV_bus_ID,rateHV,rateMV,rateLV\n')
for trans in Transformers3: #get each element out of list
    name=trans.loc_name # get name of the line
    busHV=trans.bushv.cterm.loc_name # get name of the line
    busMV=trans.busmv.cterm.loc_name # get name of the line
    busLV=trans.buslv.cterm.loc_name # get name of the line
    rateHV = trans.pRating_h.ContRating
    rateMV = trans.pRating_m.ContRating
    rateLV = trans.pRating_l.ContRating
    ftrans.write('%s, %s, %s, %s, %s, %s, %s \n' %(name,busHV,busMV,busLV,rateHV,rateMV,rateLV))
ftrans.close()

#transformer data for extracting resistance and reactance
ftrans2 = open('data/csv/transformers3_char.csv','w')
ftrans2.write('transformer_ID,rHV,xHV,rMV,xMV,rLV,xLV,baseHV,baseMV,baseLV\n')
for trans in Transformers3_char: #get each element out of list
    name=trans.loc_name # get name of the line
    rHV   = trans.r1pu_h
    xHV   = trans.x1pu_h
    rMV   = trans.r1pu_h
    xMV   = trans.x1pu_h
    rLV   = trans.r1pu_h
    xLV   = trans.x1pu_h
    baseHV = trans.strn3_h
    baseMV = trans.strn3_m
    baseLV = trans.strn3_l
    ftrans2.write('%s, %s, %s, %s, %s, %s, %s, %s, %s, %s\n' %(name,rHV,xHV,rMV,xMV,rLV,xLV,baseHV,baseMV,baseLV))
ftrans2.close()

#terminal data
fterm = open('data/csv/terminals.csv','w')
fterm.write('terminal,basekV\n')
for terminal in Terminals:
    name = terminal.loc_name
    basekv = terminal.uknom
    fterm.write('%s, %s \n' %(name,basekv))
fterm.close()


#generator data
fgen = open('data/csv/generators.csv','w')
fgen.write('generator,type,terminal,category,pmin,pmax,qmin,qmax\n')
for generator in SynGenerators:
    name = generator.loc_name
    terminal = generator.bus1.cterm.loc_name
    category = generator.aCategory
    pmin = generator.Pmin_uc
    pmax = generator.Pmax_uc
    qmin = generator.cQ_min
    qmax = generator.cQ_max
    fgen.write('%s,Sync,%s,%s,%s,%s,%s,%s \n' %(name,terminal,category,pmin,pmax,qmin,qmax))
for generator in StatGenerators:
    name = generator.loc_name
    terminal = generator.bus1.cterm.loc_name
    category = generator.aCategory
    pmin = generator.Pmin_uc
    pmax = generator.Pmax_uc
    qmin = generator.cQ_min
    qmax = generator.cQ_max
    fgen.write('%s,Stat,%s,%s,%s,%s,%s,%s \n' %(name,terminal,category,pmin,pmax,qmin,qmax))
fgen.close()

#load data
fload = open('data/csv/loads.csv','w')
fload.write('load,terminal,PD,QD\n')
for load in Loads:
    name = load.loc_name
    terminal = load.bus1.cterm.loc_name
    PD   = load.plini
    QD   = load.qlini
    fload.write('%s,%s,%s,%s \n' %(name,terminal,PD,QD))
fload.close()

#coupling switches on bus bars
fswitch = open('data/csv/busbarswitches.csv','w')
fswitch.write('name,bus_from,bus_to,status\n')
for switch in BusbarSwitches:
    name = switch.loc_name
    try:
        bus1 = switch.bus1.cterm.loc_name
    except:
        bus1='None'
    try:
        bus2 = switch.bus2.cterm.loc_name
    except:
        bus2 = 'None'
    stat = switch.on_off
    fswitch.write('%s,%s,%s,%s \n' %(name,bus1,bus2,stat))
fswitch.close()
