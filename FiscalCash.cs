/*
This file is part of PF500.

PF500 is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

PF500 is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with PF500. If not, see <http://www.gnu.org/licenses/>.

*/

////////////////////////////////////////
//
// COM communication by Boro Sitnikovski
// Protocol for Synergy PF500
//
// Revision: 24.12.2011
// Revision: 27.12.2011
// Revision: 07.02.2012
// Revision: 09.02.2012
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;

public class FiscalCash
{

    #region Variables
    bool initialized = false;
    string ComPort;

    IList<byte> opcodes = new List<byte>(); 
    StringBuilder FileIn = new StringBuilder(255);
    StringBuilder FileOut = new StringBuilder(255);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public static extern int GetShortPathName(
             [MarshalAs(UnmanagedType.LPTStr)]
                   string path,
             [MarshalAs(UnmanagedType.LPTStr)]
                   StringBuilder shortPath,
             int shortPathLength
    );

    #endregion

    #region Private Functions
    private byte[] InitRawCommand(string command)
    {
        CheckInterfaceInit();
        AppendStringToOpcodes(command);

        byte[] returnVal = StartProcessAndWait();

        try
        {
            CloseInterface();
        }
        catch (Exception) { }

        return returnVal;
    }
    private byte[] StartProcessAndWait()
    {
        File.WriteAllBytes(FileIn.ToString(), opcodes.ToArray());
        Process process = new Process();
        process.StartInfo.FileName = "pf500.exe";
        process.StartInfo.Arguments = ComPort + " " + FileIn.ToString() + " " + FileOut.ToString();
        process.Start();
        process.WaitForExit();
        return File.ReadAllBytes(FileOut.ToString());
    }
    private void CheckInterfaceInit()
    {
        if (initialized == false) throw new Exception("Ne e inicijaliziran interfejsot.");
    }
    private string ParseProductName(string ProductName)
    {
        string retval = "";
        for (int i = 0; i < ProductName.Length; i++)
        {
            switch (ProductName[i])
            {
                case 'А': retval += 'A'; break;
                case 'Б': retval += 'B'; break;
                case 'В': retval += 'V'; break;
                case 'Г': retval += 'G'; break;
                case 'Д': retval += 'D'; break;
                case 'Ѓ': retval += '\\'; break;
                case 'Е': retval += 'E'; break;
                case 'Ж': retval += '@'; break;
                case 'З': retval += 'Z'; break;
                case 'Ѕ': retval += 'Y'; break;
                case 'И': retval += 'I'; break;
                case 'Ј': retval += 'J'; break;
                case 'К': retval += 'K'; break;
                case 'Л': retval += 'L'; break;
                case 'Љ': retval += 'Q'; break;
                case 'М': retval += 'M'; break;
                case 'Н': retval += 'N'; break;
                case 'Њ': retval += 'W'; break;
                case 'О': retval += 'O'; break;
                case 'П': retval += 'P'; break;
                case 'Р': retval += 'R'; break;
                case 'С': retval += 'S'; break;
                case 'Т': retval += 'T'; break;
                case 'Ќ': retval += ']'; break;
                case 'У': retval += 'U'; break;
                case 'Ф': retval += 'F'; break;
                case 'Х': retval += 'H'; break;
                case 'Ц': retval += 'C'; break;
                case 'Ч': retval += '^'; break;
                case 'Џ': retval += 'X'; break;
                case 'Ш': retval += '['; break;
                case 'а': retval += 'a'; break;
                case 'б': retval += 'b'; break;
                case 'в': retval += 'v'; break;
                case 'г': retval += 'g'; break;
                case 'д': retval += 'd'; break;
                case 'ѓ': retval += '|'; break;
                case 'е': retval += 'e'; break;
                case 'ж': retval += '`'; break;
                case 'з': retval += 'z'; break;
                case 'ѕ': retval += 'y'; break;
                case 'и': retval += 'i'; break;
                case 'ј': retval += 'j'; break;
                case 'к': retval += 'k'; break;
                case 'л': retval += 'l'; break;
                case 'љ': retval += 'q'; break;
                case 'м': retval += 'm'; break;
                case 'н': retval += 'n'; break;
                case 'њ': retval += 'w'; break;
                case 'о': retval += 'o'; break;
                case 'п': retval += 'p'; break;
                case 'р': retval += 'r'; break;
                case 'с': retval += 's'; break;
                case 'т': retval += 't'; break;
                case 'ќ': retval += '}'; break;
                case 'у': retval += 'u'; break;
                case 'ф': retval += 'f'; break;
                case 'х': retval += 'h'; break;
                case 'ц': retval += 'c'; break;
                case 'ч': retval += '~'; break;
                case 'џ': retval += 'x'; break;
                case 'ш': retval += '{'; break;
                default: retval += ProductName[i]; break;
            }
        }
        return retval;
    }
    private void AppendStringToOpcodes(string opcode)
    {
        byte[] tmp = Encoding.ASCII.GetBytes(opcode);
        for (int i = 0; i < opcode.Length; i++) opcodes.Add(tmp[i]);
    }
    #endregion

    #region Public Functions
    ~FiscalCash()
    {
        try
        {
            CloseInterface();
        }
        catch (Exception) { }
    }

    public FiscalCash(string ComPort = "COM1")
    {
        this.ComPort = ComPort;
    }

    public void InitInterface()
    {
        if (initialized == true) throw new Exception("Interfejsot e vekje inicijaliziran.");
        initialized = true;

        GetShortPathName(@Path.GetTempFileName(), FileIn, FileIn.Capacity);
        GetShortPathName(@Path.GetTempFileName(), FileOut, FileOut.Capacity);

    }

    public void CloseInterface()
    {
        if (initialized == false) throw new Exception("Interfejsot e vekje zatvoren.");
        initialized = false;
        opcodes.Clear();
        //File.Delete(FileIn.ToString());
        File.Delete(FileOut.ToString());
    }

    public bool CheckError()
    {
        this.InitInterface();
        byte[] test = InitRawCommand("!");
        return (test[3] & 1) == 1;
    }

    public byte[] IssueBill()
    {
        return InitRawCommand("5\t\r\n8 \r\n\r\n");
    }

    public byte[] IssueStorno()
    {
        return InitRawCommand("5\t\r\nV \r\n\r\n");
    }

    public byte[] DailyFiscalClose()
    {
        return InitRawCommand("E\r\n");
    }

    public byte[] GetTimeDate()
    {
        return InitRawCommand(">\r\n");
    }

    public byte[] SetTimeDate(string time)
    {
        return InitRawCommand("=" + time + "\r\n");
    }

    public byte[] DetailedReport(string timeStart, string timeEnd)
    {
        return InitRawCommand("^" + timeStart + "," + timeEnd + "\r\n");
    }

    public byte[] ShortReport(string timeStart, string timeEnd)
    {
        return InitRawCommand("O" + timeStart + "," + timeEnd + "\r\n");
    }

    /**
     * 
     * TaxCategory:
     * 192 = А (default)
     * 193 = Б
     * 194 = В
     * 195 = Г
     * */
    public void AddProduct(string ProductName, string ProductPrice, string Quantity = "1.000" ,byte  TaxCategory = 192)
    {
        CheckInterfaceInit();
        if (opcodes.Count == 0) AppendStringToOpcodes("01,0000,1\r\n");
        AppendStringToOpcodes("1" + ParseProductName(ProductName) + "\t");
        opcodes.Add(TaxCategory);
            //0xC0);
        AppendStringToOpcodes(ProductPrice + "*" + Quantity + "\r\n");
    }

    public void AddProductStorno(string ProductName, string ProductPrice, string Quantity = "1.000",byte  TaxCategory = 192)

    {
        CheckInterfaceInit();
        if (opcodes.Count == 0) AppendStringToOpcodes("U1,0000,1\r\n");
        AppendStringToOpcodes("1" + ParseProductName(ProductName) + "\t");
        opcodes.Add(TaxCategory);
        //opcodes.Add(0xC0);
        AppendStringToOpcodes(ProductPrice + "*" + Quantity + "\r\n");
    }

    public string ConvertDateTimeToSynergyDate(DateTime value, bool dash = true)
    {
        string DD = value.Day.ToString("D2"), MM = value.Month.ToString("D2"), YY = (value.Year - 2000).ToString("D2");
        return DD + ((dash == true) ? "-" : "") + MM + ((dash == true) ? "-" : "") + YY;
    }

    #endregion
}
