if (WScript.Arguments.length < 1) {
  Trace("Usage: acquire_admin_rights.js <instance-name>");
  WScript.Quit(1);
}

var instance = WScript.Arguments(0);

Trace("Processing instance '" + instance + "'");
Trace("Setting up single user mode...");
EnableSingleUserMode(instance);
Trace("Done");
Trace("Restarting the service '" + InstanceServiceName(instance) + "'");
RestartService(InstanceServiceName(instance));
Trace("Done");
Trace("Adding '" + GetCurrentUser() + "' as SQL server admin...");
AddAdmin(instance, GetCurrentUser());
Trace("Done");
Trace("Disabling single user mode...");
DisableSingleUserMode(instance);
Trace("Done");
Trace("Restarting the service '" + InstanceServiceName(instance) + "'");
RestartService(InstanceServiceName(instance));
Trace("Done");

function AddAdmin(instance, user) {
  ExecuteCommand(instance, "EXEC sp_addsrvrolemember '" + user + "', 'sysadmin'");
}

function InstanceServiceName(instance) {
  return instance == "MSSQLSERVER" ? "MSSQLSERVER" : "MSSQL$" + instance;
}

function InstanceProperty(instance, property) {
  return "SqlServiceAdvancedProperty.PropertyIndex=13,SqlServiceType=1,PropertyName='" + property + "',ServiceName='" + instance + "'";
}

function GetPropertyValue(wmi, path) {
  return wmi.Get(path).PropertyStrValue;
}

function GetCurrentUser() {
  var wmi = GetObject("WINMGMTS:\\\\.\\root\\cimv2");
  var systems = new Enumerator(wmi.ExecQuery("SELECT * FROM Win32_ComputerSystem"));
  for (; !systems.atEnd(); systems.moveNext()) {
    return systems.item().UserName;
  }
}

function SetPropertyValue(wmi, path, value) {
  var arg = wmi.Get(path).Methods_("SetStringValue").inParameters.SpawnInstance_();
  arg.Properties_.Item("StrValue") = value;
  var result = wmi.ExecMethod(path, "SetStringValue", arg);
  if (result.ReturnValue != 0) {
    throw new Error("Failed to set property '" + path + "' to value '" + value + "'");
  }
}

function AppendSingleUserMode(paramString) {
  return "-m;" + paramString;
}

function StripSingleUserMode(paramString) {
  return paramString.substr("-m;".length);
}

function ModifyStartupParameters(instance, functor) {
  var wmi = OpenSqlWmiNamespace(instance);
  var propertyPath = InstanceProperty(instance, "STARTUPPARAMETERS");
  var startupParameters = GetPropertyValue(wmi, propertyPath);
  SetPropertyValue(wmi, propertyPath, functor(startupParameters));
}

function EnableSingleUserMode(instance) {
  ModifyStartupParameters(instance, AppendSingleUserMode);
}

function DisableSingleUserMode(instance) {
  ModifyStartupParameters(instance, StripSingleUserMode);
}

function OpenSqlWmiNamespace(instance) {
  var sqlNamespace = "WINMGMTS:\\\\.\\root\\Microsoft\\SqlServer\\";
  var wql = "SELECT * from ServerSettings WHERE InstanceName='" + instance + "'";
  var wmi = null;
  try {
    wmi = GetObject(sqlNamespace + "ComputerManagement10");
  }
  catch (exception) {
    return GetObject(sqlNamespace + "ComputerManagement");
  }
  return !(new Enumerator(wmi.ExecQuery(wql))).atEnd() ? 
    wmi : GetObject(sqlNamespace + "ComputerManagement");
}

function Connect(instance) {
  var connection = new ActiveXObject("ADODB.Connection");
  connection.ConnectionString = instance != "MSSQLSERVER" ?
    "Provider=sqloledb;Data Source=(local)\\" + instance + ";Integrated Security=SSPI;" :
    "Provider=sqloledb;Data Source=(local);Integrated Security=SSPI;";
  Trace(connection.ConnectionString);
  connection.Open();
  connection.CommandTimeout = 600;
  return connection;
}

function ExecuteCommand(instance, sql) {
  var connection = Connect(instance);
  connection.Execute(sql);
}

function GetService(name) {
  var wmi = GetObject("WINMGMTS:\\\\.\\root\\cimv2");
  var services = new Enumerator(wmi.ExecQuery("SELECT * FROM Win32_Service WHERE Name='" + name + "'"));
  for (; !services.atEnd(); services.moveNext()) {
    return services.item();
  }
}

function WaitForState(name, service, state) {
  while (service.State != state) {
    WScript.Sleep(100);
    service = GetService(name);
  }
}

function RestartService(name) {
  var service = GetService(name);
  if (service == null || service.StartMode == "Disabled")
    return;

  Trace("Stopping the service '" + name + "'");
  service.StopService();
  WaitForState(name, service, "Stopped");
  
  Trace("Starting the service '" + name + "'");
  service.StartService();
  WaitForState(name, service, "Running");
}

function Trace(message) {
  WScript.Echo(message);
}