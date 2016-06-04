using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificationISPD
{
    class WindowsReaderProgram : IReaderProgram
    {
        public List<string> ListProgram {get; set;}

        public void ReadProgram()
        {
            ListProgram.Clear();
            using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)) // или 32
            using (var uninstall = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
            {
                string[] skeys = uninstall.GetSubKeyNames(); // Все подключи из key
                int length = skeys.Length;
                // Проход по всем подключам
                for (int i = 0; i < length; i++)
                { // Получаем очередной подключ 
                    RegistryKey appKey = uninstall.OpenSubKey(skeys[i]);
                    string name;
                    try // Пробуем получить значение DisplayName 
                    {
                        name = appKey.GetValue("DisplayName").ToString();
                    }
                    catch (Exception)
                    { // Если не указано имя, то пропускаем ключ 
                        continue;
                    }
                    ListProgram.Add(name);
                    MessageBox.Show(name);
                    appKey.Close();
                }
                uninstall.Close();
            }

        }
    }
}
