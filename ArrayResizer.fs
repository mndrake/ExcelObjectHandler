namespace Utility
open System
open System.Collections.Generic
open ExcelDna.Integration

module ArrayResizer =

  let private ResizeJobs = new Queue<ExcelReference>()

  let internal EnqueueResize(caller:ExcelReference,rows:int,columns:int) =
      let target = new ExcelReference(
                        caller.RowFirst, 
                        caller.RowFirst + rows - 1, 
                        caller.ColumnFirst, 
                        caller.ColumnFirst + columns - 1, 
                        caller.SheetId)
      ResizeJobs.Enqueue(target)

  let private DoResize(target:ExcelReference) =
      try
          XlCall.Excel(XlCall.xlcEcho, false) |> ignore
           // Get the formula in the first cell of the target
          let formula:string = 
              XlCall.Excel(XlCall.xlfGetCell, 41, target) |> unbox
          let firstCell = new ExcelReference(
                                  target.RowFirst, 
                                  target.RowFirst, 
                                  target.ColumnFirst, 
                                  target.ColumnFirst, 
                                  target.SheetId)
          let isFormulaArray:bool = 
              XlCall.Excel(XlCall.xlfGetCell, 49, target) |> unbox
          if isFormulaArray then
              let oldSelectionOnActiveSheet = 
                  XlCall.Excel(XlCall.xlfSelection)
              let oldActiveCell = XlCall.Excel(XlCall.xlfActiveCell)

              // Remember old selection and select the first cell of the target
              let firstCellSheet:string = 
                  XlCall.Excel(XlCall.xlSheetNm, firstCell) |> unbox
              XlCall.Excel(XlCall.xlcWorkbookSelect,firstCellSheet) |> ignore
              let oldSelectionOnArraySheet = XlCall.Excel(XlCall.xlfSelection)
              XlCall.Excel(XlCall.xlcFormulaGoto, 
                           firstCell) |> ignore

              // Extend the selection to the whole array and clear
              XlCall.Excel(XlCall.xlcSelectSpecial, 6) |> ignore
              let oldArray:ExcelReference = 
                  XlCall.Excel(XlCall.xlfSelection) |> unbox

              oldArray.SetValue(ExcelEmpty.Value) |> ignore
              XlCall.Excel(XlCall.xlcSelect,
                           oldSelectionOnArraySheet) |> ignore
              XlCall.Excel(XlCall.xlcFormulaGoto,
                           oldSelectionOnActiveSheet) |> ignore
          // Get the formula and convert to R1C1 mode
          let isR1C1Mode:bool = 
              XlCall.Excel(XlCall.xlfGetWorkspace, 4) |> unbox |> unbox
          let mutable formulaR1C1 = formula
          if not isR1C1Mode then
              // Set the formula into the whole target
              formulaR1C1 <- XlCall.Excel(
                                 XlCall.xlfFormulaConvert, 
                                 formula, 
                                 true, 
                                 false, 
                                 ExcelMissing.Value, 
                                 firstCell) :?> string
          let ignoredResult = new System.Object()
          let  retval = XlCall.TryExcel(
                          XlCall.xlcFormulaArray, 
                          ref ignoredResult, 
                          formulaR1C1, 
                          target)
          if not (retval = XlCall.XlReturn.XlReturnSuccess) then
              firstCell.SetValue("'" + formula) |> ignore
      finally
             XlCall.Excel(XlCall.xlcEcho, true) |> ignore
   
  let private DoResizing() =
      while ResizeJobs.Count > 0 do
          DoResize(ResizeJobs.Dequeue())


  /// resizes array output of Excel UDFs
  [<ExcelFunction>]
  let Resize(array:obj[,]) =
      let caller:ExcelReference = XlCall.Excel(XlCall.xlfCaller) |> unbox
      if (caller = null) then array
      else
          let rows = array.GetLength(0)
          let columns = array.GetLength(1)

          if ((caller.RowLast - caller.RowFirst + 1 <> rows) ||
              (caller.ColumnLast - caller.ColumnFirst + 1 <> columns)) then       
                 EnqueueResize(caller, rows, columns)
                 ExcelAsyncUtil.QueueAsMacro(ExcelAction(DoResizing))
                 [|[| (box ExcelError.ExcelErrorNA) |]|] |> array2D
          else
                 array