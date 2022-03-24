from ortools.sat.python import cp_model
import pandas as pd

# Loading data
# We have extracted data from the excel file below
data = pd.read_excel('Assignment_DA_1_data.xlsx', index_col=0, sheet_name=None)
projects = list (data['Projects'].index)
months = list(data['Projects'].columns)
contractor = list(data['Quotes'].index)
values = list(data['Value']['Value'])
jobs = list(data['Quotes'].columns)


class SolutionPrinter(cp_model.CpSolverSolutionCallback):
        
        def __init__(self, projectVars, allSol,profit):
            super().__init__()
            self.projectVars = projectVars
            self.allSol = allSol
            self.profit = profit
            self.solutionCounter = 0
                
        def OnSolutionCallback(self):
            self.solutionCounter += 1
            for p in projects:
                if self.Value(self.projectVars[p]):
                    print('Taken projects: '+p+'\n')
                else: 
                    print('Available projects: '+p+'\n')
                    
                for m in months:
                    for c in contractor:
                        if self.Value(self.allSol[(p,m,c)]):
                            print(p+' of '+data['Projects'][m][p]+" in the month of "+m)
            
            print("Profit is: "+str(self.Value(self.profit)))

        
model = cp_model.CpModel()

# We need a variable which will keep tract of every possible job a contractor can do
contractorProject = {} #This contains list of jobs a contactor can do
for c in contractor:
    tmp = []
    for j in jobs:
        if pd.notna(data['Quotes'][j][c]):
            tmp.append(j)
    contractorProject[c] = tmp


# We also need all the projects in a bool variable to exastively check all solutions and setting values to true or false

allSol = {}
projectVars = {}
for p in projects:
    for m in months:
        for c in contractor:
            allSol[(p,m,c)] = model.NewBoolVar( str(p)+str(m)+str(c))
            if data['Projects'][m][p] in contractorProject[c]:
                model.Add(allSol[(p,m,c)] <=1)
            else:
                model.Add(allSol[(p,m,c)] ==0)
    projectVars[p] = model.NewBoolVar(p)
#----------------------------------- Task C---------------------------------------------

# We cannont let 2 contractors work on the same project. We need to add a constraint for it TASK C
for c in contractor: 
    for m in months:
        model.Add(sum(allSol[(p,m,c)] for p in projects) <=1)

#----------------------------------- Task D AND E---------------------------------------------

for p in projects: 
    for m in months:
        # Only 1 contactor should accept the project TASK D
        if pd.notna(data['Projects'][m][p]):
            model.Add(sum(allSol[(p,m,c)]for c in contractor) == 1).OnlyEnforceIf(projectVars[p])

        # Not taken no one should be contracted to work on it TASK E
        model.Add(sum(allSol[(p,m,c)]for c in contractor) == 0).OnlyEnforceIf(projectVars[p].Not())       

#----------------------------------- Task F---------------------------------------------
# Finding which projects are dependent on the others
dependent = {}    
for pr in projects:
    tmp = {}
    for pc in projects:
        tmp[pc] = model.NewBoolVar(pr+pc)
    dependent[pr] = tmp

for pc in projects:
    for pr in projects:
        if data['Dependencies'][pc][pr] == 'required':
            model.AddBoolAnd([dependent[pc][pr]]).OnlyEnforceIf(projectVars[pr])
        elif data['Dependencies'][pc][pr] == 'conflict':
            model.AddBoolAnd([dependent[pc][pr].Not()])
        else:
            model.AddBoolAnd([dependent[pc][pr]])

for i in projects:
    for j in projects:
        model.AddBoolAnd(projectVars[i]).OnlyEnforceIf(dependent[i][j])
#----------------------------------- Task G---------------------------------------------
t1 = []
t2 = []

for p in range(len(projects)):
    t1.append( projectVars[projects[p]]*values[p] )
    for c in range(len(contractor)):
        for m in range(len(months)):
            if (pd.notna(data['Projects'][months[m]][projects[p]])):
                if pd.notna((data['Quotes'].loc[contractor[c],data['Projects'].loc[projects[p],months[m]]])):
                    t2.append(allSol[(projects[p],months[m],contractor[c])] * int(data['Quotes'].loc[contractor[c],data['Projects'].loc[projects[p],months[m]]]))


totalVal = sum(t1)
totalCost = sum(t2)
valuation = totalVal - totalCost

model.Add(valuation >= 2160)
profit = model.NewIntVar(0,sum(values), 't')
model.Add(profit == valuation)    

solver = cp_model.CpSolver()
solution_printer = SolutionPrinter(projectVars, allSol,profit )
slve = solver.SearchForAllSolutions(model, solution_printer)
print('Solutions Count ',(solution_printer.solutionCounter))
print(solver.StatusName(slve))