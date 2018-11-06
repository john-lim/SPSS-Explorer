* Encoding: UTF-8.

SET PRINTBACK ON MPRINT ON.
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
    /FORMAT SORT BLANK(.3)
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

