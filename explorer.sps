* Encoding: UTF-8.
BEGIN PROGRAM.
import spss
import SpssClient
from Tkinter import *
import tkMessageBox
import time

selectedVar=[] # List to hold selected variables.
logs=["Log of exploratory factor analysis:"] # Log to hold the list of changes.

# Uses a dialog box used to select variables that the user wants to run in the analysis
def selectVar():
    # TO-DO: Fix the gui to use two listboxes and improve the look-and-feel of the window
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

# Standard settings for reliability analysis
def reliabilityAnalysis(sublistVar):
    spss.Submit("""
    RELIABILITY
        /VARIABLES={names}
        /SCALE('ALL VARIABLES') ALL
        /MODEL=ALPHA
        /SUMMARY=TOTAL.
    """.format(names=' '.join(sublistVar)))
    time.sleep(1)

# Gets the kmo to check if there is adequate correlation between variables.
def adequacyCheck():
    LogOutputDoc=SpssClient.GetDesignatedOutputDoc()
    factorAnalysis(selectedVar)
    LogOutputDoc.SelectLastOutput
    OutputItemList=LogOutputDoc.GetOutputItems()
    
    # Loop through output items to get the kmo table
    for index in reversed(range(OutputItemList.Size())):
        OutputItem = OutputItemList.GetItemAt(index)
        if OutputItem.GetDescription() == "KMO and Bartlett's Test":
            KMOBartletts = OutputItem.GetSpecificType()
            KMOBartlettDataCells = KMOBartletts.DataCellArray()
            kmo = (float(KMOBartlettDataCells.GetValueAt(0,0)))
            
            if kmo < 0.7:
                tkMessageBox.showerror("Error", "The variable set you specified is not adequately correlated (as indicated by the KMO: {currKmo}, which should be greater than 0.700).".format(currKmo=kmo))
            else:
                logs.append("Passed adequacy check with a KMO value: {currKmo}".format(currKmo=kmo))
                print("Passed adequacy check with a KMO value: {currKmo}".format(currKmo=kmo))

# Check for cross loadings less than 0.2 and remove them from the selected dataset
def discriminantCheck(noChangeDisc):
    global crossDict
    noChangeDisc=True
    while True:
        StageOutputDoc = SpssClient.NewOutputDoc()
        StageOutputDoc.SetAsDesignatedOutputDoc()
        factorAnalysis(selectedVar)
        OutputItemList=StageOutputDoc.GetOutputItems()
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
            OutputDocsList = SpssClient.GetOutputDocuments()
            LogOutputDoc=OutputDocsList.GetItemAt(0)
            LogOutputDoc.SetAsDesignatedOutputDoc()
            print("Removed {} because it has a low cross loading.".format(removeVar))
            factorAnalysis(selectedVar)
            StageOutputDoc.CloseDocument()
        except:
            OutputDocsList = SpssClient.GetOutputDocuments()
            LogOutputDoc=OutputDocsList.GetItemAt(0)
            LogOutputDoc.SetAsDesignatedOutputDoc()
            print("Completed check for discriminance")
            StageOutputDoc.CloseDocument()
            return noChangeDisc
            break
        
# Look for loadings below the threshold and remove them if possible
def convergentCheck(noChangeCon):
    global minDict
    noChangeCon=True
    while True:
        StageOutputDoc = SpssClient.NewOutputDoc()
        StageOutputDoc.SetAsDesignatedOutputDoc()
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
            OutputDocsList = SpssClient.GetOutputDocuments()
            LogOutputDoc=OutputDocsList.GetItemAt(0)
            LogOutputDoc.SetAsDesignatedOutputDoc()
            print("Removed {} because it has a low loading.".format(removeVar))
            factorAnalysis(selectedVar)
            StageOutputDoc.CloseDocument()
        except:
            OutputDocsList = SpssClient.GetOutputDocuments()
            LogOutputDoc=OutputDocsList.GetItemAt(0)
            LogOutputDoc.SetAsDesignatedOutputDoc()
            print("Completed check for convergence")
            StageOutputDoc.CloseDocument()
            return noChangeCon
            break

def reliabilityCheck(noChangeRel):
    noChangeRel=True
    grouping=[]
    StageOutputDoc = SpssClient.NewOutputDoc()
    StageOutputDoc.SetAsDesignatedOutputDoc()
    factorAnalysis(selectedVar)
    OutputItemList=StageOutputDoc.GetOutputItems()
    for index in range(OutputItemList.Size()):
        OutputItem = OutputItemList.GetItemAt(index)
        if OutputItem.GetDescription() == "Pattern Matrix":
            PatternMatrix = OutputItem.GetSpecificType()
            MatrixDataCells = PatternMatrix.DataCellArray()
            MatrixRowLabels = PatternMatrix.RowLabelArray()
            for j in range(MatrixDataCells.GetNumColumns()):
                    sublist=[]
                    for i in range(MatrixDataCells.GetNumRows()):
                        val=None
                        try:
                            val=float(MatrixDataCells.GetValueAt(i,j))
                            if val > 0.3:
                                sublist.append(MatrixRowLabels.GetValueAt(i,1))
                        except:
                            pass            
                    grouping.append(sublist)
    StageOutputDoc.CloseDocument()
    for group in grouping:
        if len(group) > 1:
            StageOutputDoc = SpssClient.NewOutputDoc()
            StageOutputDoc.SetAsDesignatedOutputDoc()
            reliabilityAnalysis(group)
            OutputItemList=SpssOutputDoc.GetOutputItems()
            for index in range(OutputItemList.Size()):
                OutputItem = OutputItemList.GetItemAt(index)
                if OutputItem.GetDescription() == "Reliability Statistics":
                    ReliabilityStatsTable = OutputItem.GetSpecificType()
                    ReliabilityStatsDataCells = ReliabilityStatsTable.DataCellArray()
                    cronbachs = float(ReliabilityStatsDataCells.GetValueAt(0,0))
                    for index in range(OutputItemList.Size()):
                        OutputItem = OutputItemList.GetItemAt(index)
                        if OutputItem.GetDescription() == "Item-Total Statistics":
                            ItemStatsTable = OutputItem.GetSpecificType()
                            ItemStatsDataCells = ItemStatsTable.DataCellArray()
                            ItemStatsLabels = ItemStatsTable.RowLabelArray()
                            for i in range(ItemStatsDataCells.GetNumRows()):
                                itemStat =  float(ItemStatsDataCells.GetValueAt(i,3))
                                if itemStat > cronbachs:
                                    removeVar = ItemStatsLabels.GetValueAt(i,1)  
                                    noChangeRel=False
                                    selectedVar.remove(removeVar)
                                    logs.append("Removed {} because of the cronbachs alpha test.".format(removeVar))
            OutputDocsList = SpssClient.GetOutputDocuments()
            LogOutputDoc=OutputDocsList.GetItemAt(0)
            LogOutputDoc.SetAsDesignatedOutputDoc()
            reliabilityAnalysis(group)
            StageOutputDoc.CloseDocument()
        
SpssClient.StartClient()
OutputDocsList = SpssClient.GetOutputDocuments()
for index in range(OutputDocsList.Size()):
    SpssOutputDoc=OutputDocsList.GetItemAt(index)
    SpssOutputDoc.CloseDocument()
LogOutputDoc = SpssClient.NewOutputDoc()
LogOutputDoc.SetAsDesignatedOutputDoc()

selectVar()
noChangeDisc=False
noChangeCon=False
noChangeRel=False
while (noChangeDisc is False) and (noChangeCon is False):
    adequacyCheck()
    while (noChangeDisc is False) and (noChangeCon is False):
        noChangeDisc=discriminantCheck(noChangeDisc)
        noChangeCon=convergentCheck(noChangeCon)
    noChangeRel=reliabilityCheck(noChangeRel)
LogOutputDoc.SetAsDesignatedOutputDoc()
factorAnalysis(selectedVar)
SpssClient.StopClient()
for log in logs:
    print(log)
print("End of EFA")
END PROGRAM.
