namespace ExcelObjectHandler
open ExcelDna.Integration

module TestRTD =

// class example
 
  // class we want to create an object handle for
  type Person =
    val FirstName:string
    val LastName:string
    new (firstName,lastName) = {FirstName=firstName; LastName=lastName}

  [<ExcelFunction(Name="Person.create")>]
  let createPerson (firstName, lastName) =
    new Person(firstName, lastName)
    |> XlCache.registerObject "Person"
 
  // function that uses the object handle
  [<ExcelFunction(Name="Person.getFirstName")>]
  let getPersonFirstName personHandle = 
     let p:Person = XlCache.lookupObject personHandle 
     p.FirstName

// slow function example
 
  // slow running function restated to object array input
  let myArray (p:obj[]) =
      let (n,a) = (unbox p.[0], unbox p.[1])
      System.Threading.Thread.Sleep 2000
      [| 1 .. n |]

  // creates object asynchronously and returns handle when done
  [<ExcelFunction(Name="MyArray.create")>]
  let createMyArray (n:int,a:int) =
    let args:obj[] = [| n; a |]
    XlCache.asyncRegisterObject(myArray,"MyArray",args)

 
  [<ExcelFunction(Name="MyArray.getSum")>]
  let getMyArraySum myArrayHandle =
     let a:int[] = XlCache.lookupObject myArrayHandle
     a |> Array.sum

  let mySlowFunction (p:obj[]) =
    let myString:string = unbox p.[0]
    System.Threading.Thread.Sleep 2000
    box myString

  [<ExcelFunction(Name="AsyncMySlowFunction")>]
  let asyncMySlowFunction (s:obj) =
    let args:obj[] = [| s |]
    XlCache.asyncRun(mySlowFunction,"MySlowFunction",args)