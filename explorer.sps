* Encoding: UTF-8.
BEGIN PROGRAM.
import spss
import SpssClient
from Tkinter import *
import tkMessageBox
import time

selectedVar=[] # List to hold selected variables.
logs=[] # Log to hold the list of changes.

# Uses a dialog box used to select variables that the user wants to run in the analysis
def selectVar():
    # Function to handle the submission of the dialog box
    def okButton(select, root):
        # Add variables to list that user selects
        for i in select.curselection():
            selectedVar.append(select.get(i))
        root.destroy()

    root = Tk() # Object to hold dialog box.
    # Frames are used for placement of elements in the dialog box
    frame1 = Frame(root)
    frame1.pack()

    scroll = Scrollbar(frame1, orient=VERTICAL)
    # Extended option allows multiple select of items.
    select = Listbox(frame1, selectmode = "extended", yscrollcommand=scroll.set, height=20)
    # Add all variables in the dataset to the dialog box
    for i in range(spss.GetVariableCount()):
        select.insert(END, spss.GetVariableName(i))
    scroll.config (command=select.yview)
    scroll.pack(side=RIGHT, fill=Y)
    select.pack(side=LEFT,  fill=BOTH, expand=1)

    # Place ok button below list
    frame2 = Frame(root)
    frame2.pack()
    b1 = Button(frame2,text=" OK",command=lambda: okButton(select, root))
    b1.pack(side=LEFT)

    root.mainloop()

# Standard settings for a factor analysis.
def factorAnalysis(selectedVar):
    spss.Submit("""
    FACTOR VAR={names}
        /MISSNG LISTWISE
        /ANALYSIS={names}
        /PRINT INITIAL KMO REPR EXTRACTION ROTATION
        /FORMAT SORT BLANK(.3)
        /CRITERIA=MINEIGEN(1) ITERATE(25)
        /EXTRACTION=PC
        /ROTATION=PROMAX(4)
        /METHOD=CORRELATION.
    """.format(names=' '.join(selectedVar)))
    time.sleep(1)

# Gets the kmo to check if there is adequate correlation between variables.
def adequacyCheck():
    # Output docs are opened and closed to ensure that a f
    SpssClient.StartClient()
    SpssOutputDoc=SpssClient.GetDesignatedOutputDoc()
    SpssOutputDoc.SelectAll()
    SpssOutputDoc.Delete()
    factorAnalysis(selectedVar)
    SpssOutputDoc.SelectLastOutput
    OutputItemList=SpssOutputDoc.GetOutputItems()
    
    # Loop through output items to get the kmo table
    for index in reversed(range(OutputItemList.Size())):
        OutputItem = OutputItemList.GetItemAt(index)
        if OutputItem.GetDescription() == "KMO and Bartlett's Test":
            KMOBartletts = OutputItem.GetSpecificType()
            KMOBartlettDataCells = KMOBartletts.DataCellArray()
            kmo = (float(KMOBartlettDataCells.GetValueAt(0,0)))
            
            if kmo < 0.7:
                tkMessageBox.showerror("Error", "The variable set you specified is not adequately correlated (as indicated by the KMO: {currKmo}, which should be greater than 0.700).".format(currKmo=kmo))
    SpssClient.StopClient()

# Check for cross loadings less than 0.2 and remove them from the selected dataset
def discriminantCheck(noChangeDisc):
    global crossDict
    noChangeDisc=True
    while True:
        SpssClient.StartClient()
        SpssOutputDoc=SpssClient.GetDesignatedOutputDoc()
        OutputItemList=SpssOutputDoc.GetOutputItems()
        for index in range(OutputItemList.Size()):
            OutputItem = OutputItemList.GetItemAt(index)
            if OutputItem.GetDescription() == "Pattern Matrix":
                PatternMatrix = OutputItem.GetSpecificType()
                MatrixDataCells = PatternMatrix.DataCellArray()
                MatrixRowLabels = PatternMatrix.RowLabelArray()
                crossDict = {}
                for i in range(MatrixDataCells.GetNumRows()):
                    var1=None
                    var2=None
                    for j in range(MatrixDataCells.GetNumColumns()):
                        if MatrixDataCells.GetValueAt(i,j) != '' and var1 == None:
                            var1 = float(MatrixDataCells.GetValueAt(i,j))
                        elif  MatrixDataCells.GetValueAt(i,j) != '':
                            var2 = float(MatrixDataCells.GetValueAt(i,j))
                            diff = var1 - var2
                            if diff < 0.2:
                                crossDict[i]=diff    
        try:                        
            minVal = min(crossDict, key=crossDict.get)
            removeVar = MatrixRowLabels.GetValueAt(minVal,1)
            logs.append("Removed {} because it has a low cross loading.".format(removeVar))
            noChangeDisc=False
            selectedVar.remove(removeVar)
            factorAnalysis(selectedVar)
            SpssClient.StopClient()
        except:
            SpssClient.StopClient()
            return noChangeDisc
            break
        
# Look for loadings below the threshold and remove them if possible
def convergentCheck(noChangeCon):
    global minDict
    noChangeCon=True
    while True:
        SpssClient.StartClient()
        SpssOutputDoc=SpssClient.GetDesignatedOutputDoc()
        SpssOutputDoc.SelectAll()
        SpssOutputDoc.Delete()
        factorAnalysis(selectedVar)
        OutputItemList=SpssOutputDoc.GetOutputItems()
        for index in range(OutputItemList.Size()):
            OutputItem = OutputItemList.GetItemAt(index)
            if OutputItem.GetDescription() == "Pattern Matrix":
                PatternMatrix = OutputItem.GetSpecificType()
                MatrixDataCells = PatternMatrix.DataCellArray()
                MatrixRowLabels = PatternMatrix.RowLabelArray()
                dictList=[]
                minDict={}
                for j in range(MatrixDataCells.GetNumColumns()):
                    tmpDict={}
                    for i in range(MatrixDataCells.GetNumRows()):
                        try:
                            tmpDict[(i, j)] = float(MatrixDataCells.GetValueAt(i,j))
                        except:
                            pass
                        dictList.append(tmpDict)
                    rowCnt = len(tmpDict)
                    minKey = min(tmpDict, key=tmpDict.get)
                    if rowCnt > 4:
                        if tmpDict[minKey] < 0.5:
                            minDict[minKey]=tmpDict[minKey]
                    elif rowCnt > 3:
                        if tmpDict[minKey] < 0.4:
                            minDict[minKey]=tmpDict[minKey]    
        try:
            minVal = min(minDict, key=minDict.get)
            removeVar = MatrixRowLabels.GetValueAt(minVal[0],1)   
            logs.append("Removed {} because it has a low loading.".format(removeVar))
            noChangeCon=False
            selectedVar.remove(removeVar)
            SpssClient.StopClient()   
        except:
            SpssClient.StopClient()
            return noChangeCon
            break
       
        
    
selectVar()
adequacyCheck()
noChangeDisc=False
noChangeCon=False
while (noChangeDisc is False) and (noChangeCon is False):
    noChangeDisc=discriminantCheck(noChangeDisc)
    noChangeCon=convergentCheck(noChangeCon)
for log in logs:
    print(log)
END PROGRAM.