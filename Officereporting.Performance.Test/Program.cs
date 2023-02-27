using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using el = Microsoft.Office.Interop.Excel;

namespace Officereporting.Performance.Test
{
    public class Program
    {
        static el.Application app = null;
        static void Main(string[] args)
        {
            try
            {
                new TestAction((me) => {
                    app = new el.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;
                },"Initial Application");
                
                MenuSelection();
                
            }
            catch { }
            finally
            {
                app.Quit();
            }

            Console.WriteLine("------- END -------\r\nPress any key to quite");
            Console.ReadKey();
        }

        [MenuItemAttirbue(1, "populate value to cell in loop")]
        public static void simpleActionInALoop()
        {
            TestAction action = new TestAction((me) => {
                var book = app.Workbooks.Add();
                var sheet = book.Worksheets[1];
                for (int index=0; index<10; index++) 
                {
                    new TestAction((m2) => {
                        sheet.PageSetup.CenterHeader = "sheet header";
                        sheet.PageSetup.RightHeader = "right sheet header";
                        sheet.PageSetup.LeftFooter = "left sheet header";

                    }, "set page header and footer", 1);

                    new TestAction((m2) => {
                        sheet.Range[sheet.Cells[index * 52 + index + 4, 1], sheet.Cells[index * 52 + index + 4, 12]].Merge();
                        sheet.Cells[index * 52 + index + 4, 1] = "o Whom It May Concern,";
                        sheet.Cells[index + 1, 1] = "o Whom It May Concern,";
                        
                    }, "set val", 1);
                }
            },"Root",0);

            
        }


        [MenuItemAttirbue(2, "populate value & add format to cell in loop")]
        public static void simpleAction2InALoop()
        {
            Console.WriteLine("simpleAction2InALoop");
        }

        static void MenuSelection()
        {
            Console.WriteLine("--------------------------------------------");

            Dictionary<string, MethodInfo> menu = new Dictionary<string, MethodInfo>();
            Assembly.GetEntryAssembly().ExportedTypes.First(t => t.Name == "Program").GetMethods().Where(m => m.GetCustomAttribute<MenuItemAttirbue>() != null).OrderBy(m => m.GetCustomAttribute<MenuItemAttirbue>().Index).ToList().ForEach(m =>
            {
                var cusAttr = m.GetCustomAttribute<MenuItemAttirbue>();
                Console.WriteLine($"{cusAttr.Index}. {cusAttr.Text}");
                menu.Add(cusAttr.Index.ToString(), m);
            });

            Console.WriteLine("--------------------------------------------");
            Console.Write("Type the menu item sequence num: ");

            var userSelected = menu["1"];
            Console.Clear(); 
            userSelected.Invoke(null, null);    
        }

    }

    [AttributeUsage(AttributeTargets.Method)]
    public class MenuItemAttirbue : Attribute
    {
        public int Index { get; set; }
        public string Text { get; set; }
        public MenuItemAttirbue(int index, string text)
        {
            this.Index = index;
            this.Text = text;
        }
    }
    public class TestAction
    {
        public TestAction(Action<TestAction> handle, string name, int level = 0)
        {
            Handle = handle;
            Level = level;
            Name = name;
            try
            {
                DateTime start = DateTime.Now;
                handle.Invoke(this);
                Cost = (DateTime.Now - start).TotalSeconds;

                PrintLevel();
                Console.WriteLine($"{name} cost: {Cost.ToString()}s");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
        public string Name { get; set; }
        public int Level { get; set; }


        /// <summary>
        /// Cost in seconds
        /// </summary>
        public double Cost { get; set; }

        private Action<TestAction> Handle { get; set; }

        private void PrintLevel()
        {
            for (int ind = 0; ind < Level; ind++) Console.Write("\t");
        }

    }
}
