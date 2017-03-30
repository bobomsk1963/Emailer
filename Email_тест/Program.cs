using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;
using System.IO;
using System.Text;

namespace Email_тест
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string[] mascommand = Environment.CommandLine.Split('/');
            if (mascommand.Length > 2)
            {
                bool MessageBoxShow = false;
                try
                {
                    MessageBoxShow = Int32.Parse(mascommand[1].Trim()) == 1;
                }
                catch { }  // Первый параметр обязательный 1 или 0  Включить или нет сообщение о об статусе завершения отправки почты
                           // Второй параметр обязательный Имя файла ***.xml с конфигурацией создоваемой если запустить программу без параметров
                SaveMailParam smp = Form1.DeSerial(mascommand[2].Trim(), MessageBoxShow);
                if (smp != null)
                {
                    smp.Text = smp.Text.Replace("\n", Environment.NewLine);

                    for (int i = 3; i < mascommand.Length;i++)
                    {
                        int pos = mascommand[i].IndexOf(':');
                        if (pos > -1)
                        {
                            string se = mascommand[i].Remove(0, pos+1);
                            se=se.Trim();
                            string ss = mascommand[i].Substring(0, pos);
                            //MessageBox.Show("-"+ss + "-   -"+se+"-");
                            switch (ss.Trim().ToUpper())
                            {
                                case "SM": smp.SMTPServer = se; break;// - SMTP сервер
                                case "PT": try { smp.SMTPPort = Int32.Parse(se); }
                                    catch { } break;// - Порт
                                case "FR": smp.FromMail = se; break;// - Адрес отправителя
                                case "PS": smp.FromPassword = se; break;// - Пароль
                                case "SL": try { smp.SSL = Int32.Parse(se) == 1; }
                                    catch { } break;// - SSL
                                case "DN": smp.DisplayName = se; break; // Имя адреса

                                case "ML":
                                    if (se.Length > 0)
                                    {
                                        if (smp.ToMail.Length > 0)
                                        {
                                            smp.ToMail = smp.ToMail + ";" + se;
                                        }
                                        else { smp.ToMail = se; }
                                    }
                                    break;// Добавление - Адреса получателя
                                case "MLC": smp.ToMail = se; break;// С очисткой- Адрес получателя

                                case "TP": smp.Tema = se; break;// - Заголовок отправления


                                case "TX": smp.Text = smp.Text + se;  // Добавляется текст к последней строке
                                    break;// - Текст письма
                                case "TXT": smp.Text = se + Environment.NewLine + smp.Text;  // Добавляется текст перед уже существующим
                                    break;// - Текст письма
                                case "TXN": smp.Text = smp.Text + se + Environment.NewLine;  // Добавляется текст к последней строке с переводом на следущую строку
                                    break;// - Текст письма
                                case "TXC": smp.Text = se; break;// С удалением предидущего текста - Текст письма
                                case "TXF":                      // Загрузка текста из файла
                                    try
                                    {
                                        smp.Text = File.ReadAllText(se, Encoding.GetEncoding(1251));//"windows-1251"));//"koi8-r"));
                                    }
                                    catch (Exception e)
                                    {
                                        if (MessageBoxShow)
                                        {
                                        MessageBox.Show(se + " - " + e.Message);
                                        }                                            
                                    }
                                    break;

                                case "FL":
                                    if (se.Length > 0)
                                    {
                                        if (smp.FileNames == null)
                                        {
                                            smp.FileNames = new List<string>();
                                        }
                                        smp.FileNames.Add(se);
                                    }
                                    break;// - Файл прикрепленный
                                case "FLC": if (smp.FileNames != null)   //Очистка списка файлов
                                    {
                                        smp.FileNames.Clear();
                                    }
                                    if (se.Length > 0)
                                    {
                                        if (smp.FileNames == null)
                                        {
                                            smp.FileNames = new List<string>();
                                        }
                                        smp.FileNames.Add(se);
                                    }
                                    break;

                                case "HT": try { smp.isHTML = Int32.Parse(se) == 1; }
                                    catch { } break;// - isHTML
                                default: break;
                            }
                        }
                    }

                    EventWaitHandle EvenWaitH = new AutoResetEvent(false);

                    StateMail sm = new StateMail(EvenWaitH, MessageBoxShow);
                    Form1.SendMail(smp, sm);

                    EvenWaitH.WaitOne();

                    Thread.Sleep(100);

                    if (sm.Result) 
                    {
                        Environment.Exit(0); 
                    } 
                    else 
                    {
                        Environment.Exit(11000); 
                    }

                }
                else 
                {
                Environment.Exit(11000); 
                }
            }

            Application.Run(new Form1());
        }
    }
}

/*
 * Первый параметр обязательный 1 или 0  Включить или нет сообщение о об статусе завершения отправки почты
 * Второй параметр обязательный Имя файла ***.xml с конфигурацией создоваемой если запустить программу без параметров 
 * далее может быть произвольное количество параметров
 * SM - SMTP сервер
 * PT - Порт
 * FR - Адрес отправителя
 * PS - Пароль
 * SL - SSL
 * 
 * ML - Адрес получателя
 * TP - Заголовок отправления
 * TX - Текст письма
 * FL - Файл прикрепленный
 * HT - isHTML
 */
