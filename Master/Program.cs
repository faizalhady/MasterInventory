
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