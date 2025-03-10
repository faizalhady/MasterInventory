
using Master.Services;
using System;
using System.Threading.Tasks;
using CratePilotSystemWS.Services.Email;

try
{
    //  Console.WriteLine();
    SapEamProcessorTest.Process();
    WorkcellUpdater.UpdateWorkcellColumn();
    //  Console.WriteLine();
    // var result = await EmailServices.SendCustomEmail(); // No need to create an instance
    // Console.WriteLine(result.msg);




    // Console.ReadLine(); 
}
catch (Exception e)
{
    Console.WriteLine(e.ToString());
    // Console.ReadLine();
}
// try
// {
//     SapEamProcessor.ProcessSapEam();
// }
// catch(Exception e)
// {
//     Console.WriteLine(e.ToString());
//     Console.ReadLine();
// }