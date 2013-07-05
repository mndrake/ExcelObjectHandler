namespace Example
open Utility
open ExcelDna.Integration
open System.Threading
  
module ClassExample =
    
  // class we want to create an object handle for
  type Person(firstName:string, lastName:string) =
    member val FirstName = firstName with get,set
    member val LastName = lastName with get,set
 
  // function to create object and pass it's handle to excel
  [<ExcelFunction(Name="Person.create")>]
  let createPerson (firstName, lastName) =
    new Person(firstName, lastName)
    |> XlCache.register "Person"
 
  // function that uses the object handle
  [<ExcelFunction(Name="Person.getFirstName")>]
  let getPersonFirstName personHandle = 
     let person:Person = XlCache.lookup personHandle
     person.FirstName
 
module AsyncRegisterExample =

  // slow function example
 
  // slow running function that returns array
  let myArray (n,a) = 
      Thread.Sleep 2000
      [| 0 .. n |]
      |> Array.map ((+) a)
 
  // creates object asynchronously and returns handle when done
  [<ExcelFunction(Name="MyArray.create")>]
  let createMyArray (n,a) =
    (n,a) |> XlCache.asyncRegister "MyArray" myArray

 
  [<ExcelFunction(Name="MyArray.getSum")>]
  let getMyArraySum myArrayHandle =
     let a:int[] = XlCache.lookup myArrayHandle
     a |> Array.sum

module AsyncRunExample =

    [<ExcelFunction(Name="MySlowFunction")>]
    let mySlowFunction(s:obj,waitTime:int) =
        Thread.Sleep waitTime
        s

    [<ExcelFunction(Name="AsyncMySlowFunction")>]
    let asyncMySlowFunction(s:obj,waitTime:int) = 
        (s,waitTime) 
        |> XlCache.asyncRun "mySlowFunction" mySlowFunction

module AsyncRunAndResizeExample =

    [<ExcelFunction(Name="MyBigSlowFunction")>]
    let myBigSlowFunction(rows,columns,wait:int) =
        Thread.Sleep wait
        Array2D.init rows columns (fun i j -> box (i+j))

    [<ExcelFunction(Name="AsyncResizeMyBigSlowFunction")>]
    let asyncResizeMyBigSlowFunction(rows,columns,wait) =
        (rows,columns,wait)
        |> XlCache.asyncRunAndResize "myBigSlowFunction" myBigSlowFunction