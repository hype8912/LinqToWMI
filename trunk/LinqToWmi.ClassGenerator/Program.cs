using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Management;

namespace LinqToWmi.ProtoGenerator
{
    class Program
    {
        static void Main(string[] arguments)
        {
            Console.WriteLine("WMIClassGen - Class generator for WMI objects\r\n");

            Arguments commands = null;

            try
            {
                commands = ArgumentParser.Create<Arguments>(arguments);
            }
            catch (ArgumentParseException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }

            if (commands.Wmi.ToLower() == "common")
            {
                GenerateCommonClasses(commands);
            }
            else
            {
                GenerateClass(commands);
            }

        }

        static void GenerateCommonClasses(Arguments commands)
        {
            String[] commonClasses = new String[] { "Win32_ComputerSystem", "Win32_FloppyDrive", "Win32_Keyboard", "Win32_LogicalDisk", "Win32_NetworkAdapter",
                                                    "Win32_NetworkConnection", "Win32_OperatingSystem", "Win32_PointingDevice", "Win32_Printer", "Win32_Process",
                                                    "Win32_Processor", "Win32_Product", "Win32_UserAccount" };

            foreach (String commonClass in commonClasses)
            {
                ManagementClass metaInfo = null;

                try
                {
                    WmiMetaInformation query = new WmiMetaInformation();
                    metaInfo = query.GetMetaInformation(commonClass);
                }
                catch (ManagementException ex)
                {
                    Console.WriteLine(String.Format("Error retrieving WMI information: {0}", ex.Message));
                    continue;
                }
                catch (Exception)
                {
                    Console.WriteLine(String.Format("Could not generate a class for object '{0}'", commonClass));
                    continue;
                }

                WmiFileGenerator generator = new WmiFileGenerator();
                generator.OutputLanguage = commands.Provider;
                generator.Namespace = commands.Ns;

                generator.GenerateFile(metaInfo, commonClass, commands.Out);
                Console.WriteLine(String.Format("Generated file for object '{0}'", commonClass));

                metaInfo.Dispose();
            }
        }

        static void GenerateClass(Arguments commands)
        {
            ManagementClass metaInfo = null;

            try
            {
                WmiMetaInformation query = new WmiMetaInformation();
                metaInfo = query.GetMetaInformation(commands.Wmi);
            }
            catch (ManagementException ex)
            {
                Console.WriteLine(String.Format("Error retrieving WMI information: {0}", ex.Message));
                return;
            }
            catch (Exception)
            {
                Console.WriteLine(String.Format("Could not generate a class for object '{0}'", commands.Wmi));
                return;
            }

            WmiFileGenerator generator = new WmiFileGenerator();
            generator.OutputLanguage = commands.Provider;
            generator.Namespace = commands.Ns;

            generator.GenerateFile(metaInfo, commands.Wmi, commands.Out);
            Console.WriteLine(String.Format("Generated file for object '{0}'", commands.Wmi));
        }
    }
}