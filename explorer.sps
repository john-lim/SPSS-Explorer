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
    # Function to handle the submission of the dialog box
    def okButton(listbox2, root):
        # Add the selected variables to the global list used for tests
        listbox2.selection_set(0, END)
        for i in listbox2.curselection():
            selectedVar.append(listbox2.get(i))
        root.destroy()
    
    def move_to_right(listbox1, listbox2):
        # Move selected items to the right list box and delete from the left.
         for i in listbox1.curselection():
             listbox2.insert(END, listbox1.get(i))
         for i in listbox1.curselection()[::-1]:
             listbox1.delete(listbox2.index(i))
                                  
    def move_to_left(listbox1, listbox2):
        # Move selected items back over to the first list box and delete from the right
        for i in listbox2.curselection():
            listbox1.insert(0, listbox2.get(i))
        for i in listbox2.curselection()[::-1]:
            listbox2.delete(listbox2.index(i))
            
    root = Tk() # Object to hold dialog box.
    root.title('EFA')
    
    l1 = Label(root, text="WARNING: Save all work before pressing 'Ok'", font=("Helvetica", 9)).grid(row=0, column=0, columnspan=3)

    #First frame holds a list box and scroll bar
    frame1 = Frame(root)
    scroll1 = Scrollbar(frame1, orient=VERTICAL)
    listbox1 = Listbox(frame1, selectmode = "extended", yscrollcommand=scroll1.set, height=15)
    listbox1.pack(side=LEFT,  fill=BOTH, expand=1)
    scroll1.config (command=listbox1.yview)
    scroll1.pack(side=LEFT, fill=Y)
    frame1.grid(row=1, column=0, padx=10, pady=10)
    
    # Middle frame holds buttons to transfer variables
    frame2 = Frame(root)
    one_to_right_button = Button(frame2, text='>', command=lambda: move_to_right(listbox1, listbox2))
    one_to_left_button = Button(frame2, text='<', command=lambda: move_to_left(listbox1, listbox2))
    one_to_right_button.pack(pady=2)
    one_to_left_button.pack(pady=2)
    frame2.grid(row=1, column=1, padx=10, pady=10)
    
    # Right frame holds list box and scroll bar for selected variables
    frame3 = Frame(root)
    scroll2 = Scrollbar(frame3, orient=VERTICAL)
    listbox2 = Listbox(frame3, selectmode = "extended", yscrollcommand=scroll2.set, height=15)
    listbox2.pack(side=LEFT,  fill=BOTH, expand=1)
    scroll2.config (command=listbox2.yview)
    scroll2.pack(side=LEFT, fill=Y)
    frame3.grid(row=1, column=2, padx=10, pady=10)
    
    # Add all variables in the dataset to the dialog box
    for i in range(spss.GetVariableCount()):
        listbox1.insert(END, spss.GetVariableName(i))

    # Place ok and cancel button below list
    frame4 = Frame(root)
    frame4.pack(padx=10, pady=10)
    b1 = Button(frame4,text=" OK",command=lambda: okButton(listbox2, root), padx=5)
    b1.pack(side=LEFT, padx=5)
    b2 = Button(frame4, text="Cancel", command=root.destroy, padx=5)
    b2.pack(side=RIGHT, padx=5)
    frame4.grid(row=2, column=0, columnspan=3, padx=10, pady=10)

    # Refreshes changes to window
    root.mainloop()

# Standard settings for a factor analysis.
def factorAnalysis(selectedVar):
    # SPSS command for dimension reduction, found under the analyze tab.
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
    time.sleep(0.5) # Must slow the program down or it will fall behind spss

# Standard settings for reliability analysis
def reliabilityAnalysis(sublistVar):
    # SPSS command for reliability analysis
    spss.Submit("""
    RELIABILITY
        /VARIABLES={names}
        /SCALE('ALL VARIABLES') ALL
        /MODEL=ALPHA
        /SUMMARY=TOTAL.
    """.format(names=' '.join(sublistVar)))
    time.sleep(0.5)

# Gets the kmo to check if there is adequate correlation between variables.
def adequacyCheck():
    # The Log Output Doc holds a running record of all the tests performed
    LogOutputDoc=SpssClient.GetDesignatedOutputDoc()
    # Calls function to run dimension reduction
    factorAnalysis(selectedVar)
    
    # Loop through output items to get the kmo table
    OutputItemList=LogOutputDoc.GetOutputItems()
    for index in reversed(range(OutputItemList.Size())):
        OutputItem = OutputItemList.GetItemAt(index)
        if OutputItem.GetDescription() == "KMO and Bartlett's Test":
            KMOBartletts = OutputItem.GetSpecificType()
            # Get the KMO pivot table
            KMOBartlettDataCells = KMOBartletts.DataCellArray()
            kmo = (float(KMOBartlettDataCells.GetValueAt(0,0)))
            
            if kmo < 0.7:
                tkMessageBox.showerror("Error", "The variable set you specified is not adequately correlated (as indicated by the KMO: {currKmo}, which should be greater than 0.700).".format(currKmo=kmo))
            else:
                logs.append("Passed adequacy check with a KMO value: {currKmo}".format(currKmo=kmo))
                print("Passed adequacy check with a KMO value: {currKmo}".format(currKmo=kmo))

# Check for cross loadings less than 0.2 and remove them from the selected dataset
def discriminantCheck(noChangeDisc):
    global crossDict # A dictionary of the variables with cross loadings
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
            logs.append("Removed {} because it has a high cross loading.".format(removeVar))
            noChangeDisc=False
            selectedVar.remove(removeVar)
            OutputDocsList = SpssClient.GetOutputDocuments()
            LogOutputDoc=OutputDocsList.GetItemAt(0)
            LogOutputDoc.SetAsDesignatedOutputDoc()
            print("Removed {} because it has a high cross loading.".format(removeVar))
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
    for group in grouping:
            sublist=[]
            for g in group:
                sublist.append(''.join([i for i in g if not i.isdigit()]))
            if len(set(sublist)) > 1:
                factor = grouping.index(group) + 1
                logs.append("Factor {} contains measures from multiple constructs. Consider running a seperate EFA with those items to discriminate between unique factors.".format(factor))

selectVar()
SpssClient.StartClient()
# Close any open output documents because they will interfere with the analysis
OutputDocsList = SpssClient.GetOutputDocuments()
for index in range(OutputDocsList.Size()):
    SpssOutputDoc=OutputDocsList.GetItemAt(index)
    SpssOutputDoc.CloseDocument()
# Create the output doc to keep a running record of the tests and changes
LogOutputDoc = SpssClient.NewOutputDoc()
LogOutputDoc.SetAsDesignatedOutputDoc()

# These booleans are used to control the conditions where the tests are run and rerun.
noChangeDisc=False
noChangeCon=False
noChangeRel=False
# Repeat all tests until there are no changes made in any of the three tests.
while (noChangeDisc is False) and (noChangeCon is False) and (noChangeRel is False):
    adequacyCheck() # Exits the program if it has a low kmo
    # Repeat discriminant and convergent check until no changes are made
    while (noChangeDisc is False) and (noChangeCon is False):
        noChangeDisc=discriminantCheck(noChangeDisc)
        noChangeCon=convergentCheck(noChangeCon)
    # Run the reliability test
    noChangeRel=reliabilityCheck(noChangeRel)
# Run the dimension reduction one final time for the Log Output Doc
LogOutputDoc.SetAsDesignatedOutputDoc()
factorAnalysis(selectedVar)
SpssClient.StopClient()
# Print a list of all the changes made during the EFA
for log in logs:
    print(log)
print("End of EFA")
END PROGRAM.