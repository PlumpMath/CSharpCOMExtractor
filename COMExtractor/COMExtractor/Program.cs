using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Runtime.InteropServices;

namespace COMExtractor
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            //COMNamespace ns = new COMNamespace("COM_MSMQ");
            //ns.NewClass("Message");
            //ns.classes["Message"].createGUIDs();
            //ns.classes["Message"].ProgId = "COM_MSMQ.Message";
            //ns.classes["Message"].NewMethod("Message");
            //ns.classes["Message"].NewMethod("ToString");
            //
            //System.IO.File.WriteAllLines("D:\\Users\\Robert\\Desktop\\Temp\\COM_MSMQ.cs", ns.ToString().Split('\n'));

            COMNamespace sock = new COMNamespace("COM_Sockets");
            COMClass c = sock.NewClass("COMSocket");
            COMMethod m = c.NewMethod("Accept");
            m.returnType = "COMSocket";
            m = c.NewMethod("Connect");
            m.NewArgument("string", "IPAddress");
            m.NewArgument("int", "port");
            m = c.NewMethod("Disconnect");
            m = c.NewMethod("Listen");
            m.NewArgument("string", "IPAddress");
            m.NewArgument("int", "port");
            m.NewArgument("int", "backlog");
            m = c.NewMethod("Receive");
            m.NewArgument("byte[]", "buffer");
            m.returnType = "int";
            m = c.NewMethod("Send");
            m.returnType = "int";
            m.NewArgument("byte[]", "buffer");
            COMProperty p = c.NewProperty("Connected");
            p.returnType = "bool";
            p = c.NewProperty("Available");
            p.returnType = "int";

            System.IO.File.WriteAllLines("D:\\Users\\Robert\\Desktop\\Temp\\COM_Sockets.cs", sock.ToString().Split('\n'));
        }
    }

    class COMPart
    {
        public string Name;
        public COMPart(string Name)
        {
            this.Name = Name;
        }
        public COMPart()
        {

        }
    }

    class COMNamespace : COMPart
    {
        public System.Collections.Generic.Dictionary<string, COMClass> classes;

        public COMNamespace(string Name)
            : base(Name)
        {
            classes = new Dictionary<string, COMClass>();
        }

        public COMClass NewClass(string Name)
        {
            COMClass c = new COMClass(Name);
            c.createGUIDs();
            classes.Add(Name, c);
            return c;
        }

        public override string ToString()
        {
            string o = "";
            o += "using System;\n";
            o += "using System.Runtime.InteropServices;\n";
            o += "\n";
            o += "namespace " + Name + "\n";
            o += "{\n";

            foreach (COMClass c in classes.Values)
            {
                o += c.ToString();
            }

            o += "}\n";
            return o;
        }
    }

    class COMClass : COMPart
    {
        public string ClassGUID;
        public string InterfaceGUID;
        public string EventsGUID;
        public string ClassRegKey
        {
            get
            {
                return "CLSID\\\\{" + ClassGUID + "}\\\\TypeLib";
            }
        }
        public string ProgId;

        public System.Collections.Generic.Dictionary<string, COMMethod> methods;
        public System.Collections.Generic.Dictionary<string, COMProperty> properties;
        public System.Collections.Generic.Dictionary<string, COMEvent> events;
        public COMClass(string Name) : base(Name)
        {
            methods = new Dictionary<string, COMMethod>();
            properties = new Dictionary<string, COMProperty>();
            events = new Dictionary<string, COMEvent>();
        }

        public void createGUIDs()
        {
            this.ClassGUID = Guid.NewGuid().ToString();
            this.InterfaceGUID = Guid.NewGuid().ToString();
            this.EventsGUID = Guid.NewGuid().ToString();
        }

        public COMMethod NewMethod(string Name)
        {
            COMMethod m = new COMMethod(Name);
            methods.Add(Name, m);
            return m;
        }

        public COMProperty NewProperty(string Name)
        {
            COMProperty p = new COMProperty(Name);
            properties.Add(Name,p);
            return p;
        }

        public COMEvent NewEvent(string Name)
        {
            COMEvent e = new COMEvent(Name);
            events.Add(Name, e);
            return e;
        }

        public override string ToString()
        {
            string o = "";
            int i = 0;

            o += "[Guid(ClassId),\n";
            o += "ProgId(ProgId),\n";
            o += "ClassInterface(ClassInterfaceType.None),\n";
            o += "ComSourceInterfaces(typeof(I" + Name + "Events))]\n";
            o += "public class " + Name + " : I" + Name + "\n";
            o += "{\n";
            o += "#region properties\n";
            foreach (COMProperty p in properties.Values)
            {
                o += "\n";
                o += p.ToString();
            }
            o += "\n";
            o += "#endregion\n";
            o += "\n";
            o += "#region events\n";
            foreach (COMEvent e in events.Values)
            {
                o += "\n";
                o += e.ToString();
            }
            o += "\n";
            o += "#endregion\n";
            o += "\n";
            o += "#region methods\n";
            foreach (COMMethod m in methods.Values)
            {
                o += "\n";
                o += m.ToString();
            }
            o += "\n";
            o += "#endregion\n";
            o += "\n";
            o += "#region GUIDs and registration\n";
            o += "public const string ClassId = \"" + ClassGUID + "\";\n";
            o += "public const string InterfaceId = \"" + InterfaceGUID + "\";\n";
            o += "public const string EventsId = \"" + EventsGUID + "\";\n";
            o += "\n";
            o += "public const string ClassRegKey = \"" + ClassRegKey + "\";\n";
            o += "public const string ProgId = \"" + ProgId + "\";\n";
            o += "[ComRegisterFunctionAttribute]\n";
            o += "public static void RegisterFunction(Type t)\n";
            o += "{\n";
            o += "try\n";
            o += "{\n";
            o += "Microsoft.Win32.RegistryKey oKey;\n";
            o += "GuidAttribute oGuid = new GuidAttribute(ClassId);\n";
            o += "\n";
            o += "oKey = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(ClassRegKey);\n";
            o += "Console.WriteLine(\"Com " + Name + ".RegisterFunction: ClassRegKey Added\");\n";
            o += "\n";
            o += "oGuid = (GuidAttribute)GuidAttribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly(), oGuid.GetType());\n";
            o += "\n";
            o += "oKey.SetValue(\"\", \"{\"+oGuid.Value+\"}\");\n";
            o += "Console.WriteLine(\"Com " + Name + ".RegisterFunction: Key Value Added\");\n";
            o += "}\n";
            o += "catch (Exception ex)\n";
            o += "{\n";
            o += "Console.WriteLine(\"Com " + Name + ".RegisterFunction Exception: \" + ex.Message);\n";
            o += "}\n";
            o += "}\n";
            o += "[ComUnregisterFunctionAttribute]\n";
            o += "public static void UnregisterFunction(Type t)\n";
            o += "{\n";
            o += "try\n";
            o += "{\n";
            o += "Microsoft.Win32.Registry.ClassesRoot.DeleteSubKey(ClassRegKey);\n";
            o += "Console.WriteLine(\"Com " + Name + ".UnregisterFunction: ClassRegKey Deleted\");\n";
            o += "}\n";
            o += "catch (Exception ex)\n";
            o += "{\n";
            o += "Console.WriteLine(\"Com " + Name + ".UnregisterFunction Exception: \" + ex.Message);\n";
            o += "}\n";
            o += "}\n";

            o += "#endregion\n";
            o += "}\n";

            o += "\n";

            o += "[Guid(\"" + InterfaceGUID + "\"),\n";
            o += "InterfaceType(ComInterfaceType.InterfaceIsDual)]\n";
            o += "public interface I" + Name + "\n";
            o += "{\n";
            i = 1;
            o += "\n";
            o += "#region properties";
            foreach (COMProperty p in properties.Values)
            {
                o += "[DispId(" + i.ToString() + ")]\n";
                o += p.ToInterfaceDecl();
                o += "\n";
                i++;
            }
            o += "#endregion\n";
            o += "\n";
            o += "#region methods\n";
            foreach (COMMethod m in methods.Values)
            {
                o += "[DispId(" + i.ToString() + ")]\n";
                o += m.ToInterfaceDecl();
                o += "\n";
                i++;
            }
            o += "#endregion\n";
            o += "\n";
            o += "}\n";

            o += "[Guid(\"" + EventsGUID + "\"),\n";
            o += "InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]\n";
            o += "public interface I" + Name + "Events\n";
            o += "{\n";
            o += "#region events\n";
            i = 1;
            foreach (COMEvent e in events.Values)
            {
                o += "[DispId(" + i.ToString() + ")]\n";
                o += e.ToInterfaceDecl();
                o += "\n";
                i++;
            }
            o += "#endregion\n";
            o += "}\n";

            return o;
        }
    }

    class COMMethod : COMPart
    {
        public struct Argument
        {
            public string argType;
            public string argName;
        }
        public System.Collections.Generic.Dictionary<string, Argument> arguments;
        public string returnType;

        public COMMethod(string Name) : base(Name)
        {
            arguments = new Dictionary<string, Argument>();
            returnType = "void";
        }

        public void NewArgument(string type, string name)
        {
            Argument a = new Argument();
            a.argType = type;
            a.argName = name;
            arguments.Add(name,a);
        }

        public override string ToString()
        {
            string o = ToStringCommon();
            o += "{\n";
            o += "throw new NotImplementedException();\n";
            o += "}\n";
            return o;
        }

        public string ToInterfaceDecl()
        {
            string o  = ToStringCommon();
            o = o.TrimEnd('\n');
            o += ";\n";
            // remove "public " from the start
            o = o.Substring(7);
            return o;
        }

        private string ToStringCommon()
        {
            string o = "";

            if (returnType.Trim() == "")
                returnType = "void";

            if (arguments.Count > 0)
            {
                o += "public " + returnType + " " + Name + "(\n";
                foreach (Argument a in arguments.Values)
                {
                    o += a.argType + " " + a.argName + ",\n";
                }
                o = o.Substring(0, o.Length - 2) + ")\n";
            }
            else
            {
                o += "public " + returnType + " " + Name + "()\n";
            }

            return o;
        }
    }

    class COMProperty : COMPart
    {
        public struct Argument
        {
            public string argType;
            public string argName;
        }
        public System.Collections.Generic.Dictionary<string, Argument> arguments;
        public string returnType;

        public COMProperty(string Name) : base(Name)
        {
            arguments = new Dictionary<string, Argument>();
            returnType = "void";
        }

        public void NewArgument(string type, string name)
        {
            Argument a = new Argument();
            a.argType = type;
            a.argName = name;
            arguments.Add(name,a);
        }

        public override string ToString()
        {
            string o = ToStringCommon();
            o += "{\n";
            o += "get { throw new NotImplementedException(); }\n";
            o += "set { throw new NotImplementedException(); }\n";
            o += "}\n";
            return o;
        }

        public string ToInterfaceDecl()
        {
            string o  = ToStringCommon();
            o = o.TrimEnd('\n');
            o += "{ get; set; }";
            // remove "public " from the start
            o = o.Substring(7);
            return o;
        }

        private string ToStringCommon()
        {
            string o = "";

            if (returnType.Trim() == "")
                returnType = "void";

            if (arguments.Count > 3)
            {
                o += "public " + returnType + " " + Name + "(\n";
                foreach (Argument a in arguments.Values)
                {
                    o += a.argType + " " + a.argName + ",\n";
                }
                o = o.Substring(0, o.Length - 2) + ")\n";
            } else {
                if (arguments.Count > 0)
                {
                    o+= "public " + returnType + " " + Name + "(";
                    foreach (Argument a in arguments.Values)
                    {
                        o += a.argType + " " + a.argName + ", ";
                    }
                    o = o.Substring(0, o.Length - 2) + ")\n";
                } else {
                    o += "public " + returnType + " " + Name + "\n";
                }
            }

            return o;
        }
    }

    class COMEvent : COMPart
    {
        public struct Argument
        {
            public string argType;
            public string argName;
        }
        public System.Collections.Generic.Dictionary<string, Argument> arguments;
        public string returnType;

        public COMEvent(string Name) : base(Name)
        {
            arguments = new Dictionary<string, Argument>();
            returnType = "void";
        }

        public void NewArgument(string type, string name)
        {
            Argument a = new Argument();
            a.argType = type;
            a.argName = name;
            arguments.Add(name,a);
        }

        public override string ToString()
        {
            string o = "";
            o += "public delegate " + returnType + " " + Name + "Delegate(";
            if (arguments.Count > 3)
            {
                o += "public " + returnType + " " + Name + "(\n";
                foreach (Argument a in arguments.Values)
                {
                    o += a.argType + " " + a.argName + ",\n";
                }
                o = o.TrimEnd(',') + ");\n";
            }
            else
            {
                if (arguments.Count > 0)
                {
                    o += "public " + returnType + " " + Name + "(";
                    foreach (Argument a in arguments.Values)
                    {
                        o += a.argType + " " + a.argName + ", ";
                    }
                    o = o.TrimEnd(',') + ");\n";
                }
                else
                {
                    o += "public " + returnType + " " + Name + "();\n";
                }
            }
            o += "public event " + Name + "Delegate " + Name + ";";
            return o;
        }

        public string ToInterfaceDecl()
        {
            string o = "event " + Name + "Delegate " + Name + ";";
            return o;
        }        
    }
}
