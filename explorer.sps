* Encoding: UTF-8.
BEGIN PROGRAM.
import spss
import SpssClient
from Tkinter import *
import tkMessageBox

root = Tk()
selectedVar=[]

def okButton(select):
    for i in select.curselection():
        selectedVar.append(select.get(i))
    root.destroy()

def selectVar():
    frame1 = Frame(root)
    frame1.pack()

    scroll = Scrollbar(frame1, orient=VERTICAL)
    select = Listbox(frame1, selectmode = "extended", yscrollcommand=scroll.set, height=6)
    for i in range(spss.GetVariableCount()):
        select.insert(END, spss.GetVariableName(i))
    scroll.config (command=select.yview)
    scroll.pack(side=RIGHT, fill=Y)
    select.pack(side=LEFT,  fill=BOTH, expand=1)

    frame2 = Frame(root)
    frame2.pack()
    b1 = Button(frame2,text=" OK",command=lambda: okButton(select))
    b1.pack(side=LEFT)

    root.mainloop()

selectVar()
spss.Submit("""
FACTOR VAR={names}
    /MISSNG LISTWISE
    /ANALYSIS={names}
    /PRINT INITIAL KMO REPR EXTRACTION ROTATION
    /FORMAT SORT BLANK(.6)
    /CRITERIA=MINEIGEN(1) ITERATE(25)
    /EXTRACTION=PC
    /ROTATION=PROMAX(4)
    /METHOD=CORRELATION.
""".format(names=' '.join(selectedVar)))

SpssClient.StartClient()
SpssOutputDoc=SpssClient.GetDesignatedOutputDoc()
OutputItemList=SpssOutputDoc.GetOutputItems()
kmoList=[]

for index in range(OutputItemList.Size()):
   OutputItem = OutputItemList.GetItemAt(index)
   if OutputItem.GetDescription() == "KMO and Bartlett's Test":
       PivotTable = OutputItem.GetSpecificType()
       val = float(DataCells.GetValueAt(0,0))
       kmoList.append(val)

currKmo= kmoList[-1]
if currKmo < 0.7:
    tkMessageBox.showinfo("Error", "The variable set you specified is not adequately correlated (as indicated by the KMO: {kmo}, which should be greater than 0.700).".format(kmo=currKmo))

SpssClient.StopClient()

END PROGRAM.






BEGIN PROGRAM.
import spss
import SpssClient

varNames=[]

for i in range(spss.GetVariableCount()):
    if spss.GetVariableType(i) == 0:
        varNames.append(spss.GetVariableName(i))

spss.Submit("""
FACTOR VAR={names}
    /MISSNG LISTWISE
    /ANALYSIS={names}
    /PRINT INITIAL KMO REPR EXTRACTION ROTATION
    /FORMAT SORT BLANK(.6)
    /CRITERIA=MINEIGEN(1) ITERATE(25)
    /EXTRACTION=PC
    /ROTATION=PROMAX(4)
    /METHOD=CORRELATION.
""".format(names=' '.join(varNames)))

SpssClient.StartClient()
SpssOutputDoc=SpssClient.GetDesignatedOutputDoc()
OutputItemList=SpssOutputDoc.GetOutputItems()
for index in range(OutputItemList.Size()):
   OutputItem = OutputItemList.GetItemAt(index)
   if OutputItem.GetDescription() == "KMO and Bartlett's Test":
       PivotTable = OutputItem.GetSpecificType()
       val = float(DataCells.GetValueAt(0,0))

      if OutputItem.GetDescription() == "Pattern Matrix":
       PivotTable = OutputItem.GetSpecificType()
       DataCells = PivotTable.DataCellArray()
       for i in range(DataCells.GetNumRows()):
           for j in range(DataCells.GetNumColumns()):
              try:
                 val = float(DataCells.GetValueAt(i,j))
                 if  val > 0:
                    print(val)
              except:
                 pass
SpssClient.StopClient()
END PROGRAM.


 * RowLabels = PivotTable.RowLabelArray()
for i in range(RowLabels.GetNumRows()):
 * Add dialogue with list of variables they can choose from
Discriminance (Are the factors far apart/separated) Are there meaningful cross loadings? Not more than 0.2 different from primary loading
1 Adequacy
1 Discriminance
2 Convergence
3 Adequacy
4 Relaibility
