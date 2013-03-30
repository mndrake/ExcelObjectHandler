namespace ExcelObjectHandler
open System.Collections.Generic
open ExcelDna.Integration
open ExcelDna.Integration.Rtd 
open ExcelDna.Integration

type public ExcelInterface() = 
  interface IExcelAddIn with
    member x.AutoOpen() = ExcelAsyncUtil.Initialize()
    member x.AutoClose() = ExcelAsyncUtil.Uninitialize()

module XlCache =
  [<Literal>]
  let RTDServer = "ExcelObjectHandler.StaticRTD"
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
  let inline lookupObject (handle:string):'a =
      lookup handle |> unbox
  let inline registerObject tag (o:obj) = 
      o |> register tag
        |> fun name -> XlCall.RTD(RTDServer, null, name)
  let inline asyncRun (myFunction, tag, myInputs:obj[]) =
      ExcelAsyncUtil.Run(tag, myInputs, fun () -> (myFunction myInputs) |> box)
      |> unbox
      |> fun result ->
            if result = box ExcelError.ExcelErrorNA then 
                box ExcelError.ExcelErrorGettingData
            else result
  let inline asyncRegisterObject (myFunction, tag, myInputs:obj[]) =
      ExcelAsyncUtil.Run(tag, myInputs, fun () -> (myFunction myInputs) |> box)
      |> unbox
      |> fun result ->
            if result = box ExcelError.ExcelErrorNA then 
                box ExcelError.ExcelErrorGettingData
            else result |> register tag
                        |> fun name -> XlCall.RTD(RTDServer, null, name)    


type public StaticRTD() =
  inherit ExcelRtdServer() 
  let _topics = new Dictionary<ExcelRtdServer.Topic, string>()
  override x.ConnectData(topic, topicInfo, newValues) =
          let name = topicInfo.[0]
          _topics.[topic] <- name
          box name
  override x.DisconnectData(topic:ExcelRtdServer.Topic) =
             _topics.[topic] |> XlCache.unregister
             _topics.Remove(topic) |> ignore

