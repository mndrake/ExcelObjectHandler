namespace Utility
open System.Collections.Generic
open Microsoft.FSharp.Reflection
open ExcelDna.Integration
open ExcelDna.Integration.Rtd

module Tuple =
    /// Takes n-tuple function f with input args and returns array function f' and array input p'
    let inline ConvertToArrayFunc (f:'T -> 'a) (args:'T) =
        let getTuple p = FSharpValue.MakeTuple(p, typeof<'T>)
        let p' = FSharpValue.GetTupleFields args
        let f' = fun p -> f(unbox <| FSharpValue.MakeTuple(p, typeof<'T>))
        (f',p')

module XlCacheUtility =
  [<Literal>]
  let RTDServer = "Utility.StaticRTD"
  let private _objects = new Dictionary<string,obj>()
  let private _tags = new Dictionary<string, int>()
  let lookup handle = 
      match _objects.ContainsKey(handle) with
      |true -> _objects.[handle]
      |false -> failwith "handle not found"
  let register tag (o:obj) = 
      let counter = 
          match _tags.ContainsKey(tag) with
          |true  -> _tags.[tag] + 1
          |false -> 1
      _tags.[tag] <- counter
      let handle = sprintf "[%s.%i]" tag counter
      _objects.[handle] <- o
      handle
  let unregister handle =    
      if _objects.ContainsKey(handle) then
          _objects.Remove(handle) |> ignore

type public StaticRTD() =
  inherit ExcelRtdServer() 
  let _topics = new Dictionary<ExcelRtdServer.Topic, string>()
  override x.ConnectData(topic, topicInfo, newValues) =
          let name = topicInfo.[0]
          _topics.[topic] <- name
          box name
  override x.DisconnectData(topic:ExcelRtdServer.Topic) =
             _topics.[topic] |> XlCacheUtility.unregister
             _topics.Remove(topic) |> ignore

module XlCache = 
  let inline lookup handle = XlCacheUtility.lookup handle |> unbox

  let inline register tag (o:obj) = 
      o |> XlCacheUtility.register tag
        |> fun name -> XlCall.RTD(XlCacheUtility.RTDServer, null, name)

  let inline asyncRun tag func args =
      let (f,p) = Tuple.ConvertToArrayFunc func args
      ExcelAsyncUtil.Run(tag, p, fun () -> box <| f p)
      |> function
         |result when result = box ExcelError.ExcelErrorNA -> box ExcelError.ExcelErrorGettingData
         |result -> result

  let inline asyncRegister tag func args =
      let (f,p) = Tuple.ConvertToArrayFunc func args
      ExcelAsyncUtil.Run(tag, p, fun () -> box <| f p)
      |> function
         |result when result = box ExcelError.ExcelErrorNA -> box ExcelError.ExcelErrorGettingData
         |result -> XlCacheUtility.register tag result
                    |> fun name -> XlCall.RTD(XlCacheUtility.RTDServer, null, name)

  let inline asyncRunAndResize tag func args =
      let (f,p) = Tuple.ConvertToArrayFunc func args
      ExcelAsyncUtil.Run(tag, p, fun () -> box <| f p)
      |> function
         |result when result = box ExcelError.ExcelErrorNA -> array2D [[box ExcelError.ExcelErrorGettingData]]
         |result -> result :?> obj[,] |> ArrayResizer.Resize 