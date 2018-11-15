* Encoding: UTF-8.
BEGIN PROGRAM.
import spss
import SpssClient
from Tkinter import *
import tkMessageBox

root = Tk() # Object to hold dialog box.
selectedVar=[] # Global List to hold variables that user selects.

def okButton(select):
    # Add variables to list that user selects
    for i in select.curselection():
        selectedVar.append(select.get(i))
    root.destroy()

# Uses a dialog box used to select variables that the user wants to run in the analysis
def selectVar():
    frame1 = Frame(root)
    frame1.pack()

    scroll = Scrollbar(frame1, orient=VERTICAL)
    select = Listbox(frame1, selectmode = "extended", yscrollcommand=scroll.set, height=20)
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

def factorAnalysis(selectedVar):
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

selectVar()
factorAnalysis(selectedVar)


SpssClient.StartClient()
SpssOutputDoc=SpssClient.GetDesignatedOutputDoc()
OutputItemList=SpssOutputDoc.GetOutputItems()

for index in range(OutputItemList.Size()):
    OutputItem = OutputItemList.GetItemAt(index)
    if OutputItem.GetDescription() == "KMO and Bartlett's Test":
       KMOBartletts = OutputItem.GetSpecificType()
       KMOBartlettDataCells = KMOBartletts.DataCellArray()
       kmo = float(DataCells.GetValueAt(0,0))
    elif OutputItem.GetDescription() == "Pattern Matrix":
        PatternMatrix = OutputItem.GetSpecificType()
        MatrixDataCells = PatternMatrix.DataCellArray()

# Adequacy check
if kmo < 0.7:
    tkMessageBox.showinfo("Error", "The variable set you specified is not adequately correlated (as indicated by the KMO: {kmo}, which should be greater than 0.700).".format(kmo=currKmo))

# Discriminant validity check
for i in range(MatrixDataCells.GetNumColumns()):
    for j in range(MatrixDataCells.GetNumRows()):
        print(DataCells.GetValueAt(i,j))
# Convergent validity check

# Reliability check

SpssClient.StopClient()

END PROGRAM.

 * RowLabels = PivotTable.RowLabelArray()
