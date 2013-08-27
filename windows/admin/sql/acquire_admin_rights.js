var instance = "MSSQLSERVER";

if (WScript.Arguments.length > 1) {
  Trace("Usage: acquire_admin_rights.js [<instance-name>]");
  WScript.Quit(1);
}
else if (WScript.Arguments.length == 1) {
  instance = WScript.Arguments(0);
}

AcquireAdminRights(instance);

function AcquireAdminRights(instance) {
  try {
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
  }
  catch (exception) {
    Trace(exception.message);
  }
}

function AddAdmin(instance, user) {
  ExecuteCommand(instance, "EXEC sp_addsrvrolemember '" + user + "', 'sysadmin'");
}

function InstanceServiceName(instance) {
  return instance == "MSSQLSERVER" ? "MSSQLSERVER" : "MSSQL$" + instance;
}

function InstanceProperty(instance, property) {
  return "SqlServiceAdvancedProperty.PropertyIndex=13,SqlServiceType=1,PropertyName='" 
    + property + "',ServiceName='" + InstanceServiceName(instance) + "'";
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

function LookupInstanceContext(instance, scope) {
  try {
    var wmi = GetObject("WINMGMTS:\\\\.\\root\\Microsoft\\SqlServer\\" + scope);
    var settings = new Enumerator(wmi.ExecQuery("SELECT * FROM ServerSettings WHERE InstanceName='" + instance + "'"));
    if (!settings.atEnd()) {
      return wmi;
    }
  }
  catch (exception) {}
  return null;
}

function EnumerateSqlNamespaces() {
  var wmi = GetObject("WINMGMTS:\\\\.\\root\\Microsoft\\SqlServer");
  return new Enumerator(wmi.ExecQuery("SELECT * FROM __NAMESPACE WHERE Name LIKE 'ComputerManagement%'"));
}

function OpenSqlWmiNamespace(instance) {
  for (var namespaces = EnumerateSqlNamespaces(); !namespaces.atEnd(); namespaces.moveNext()) {
    var wmi = LookupInstanceContext(instance, namespaces.item().Name);
    if (wmi != null) {
      return wmi;
    }
  }

  throw new Error("Instance '" + instance + "' not found.");
}

function ConnectionStringInstanceComponent(instance) {
  return instance != "MSSQLSERVER" ? "\\" + instance : "";
}

function ConnectionString(instance) {
  return "Provider=sqloledb;Data Source=(local)" + 
    ConnectionStringInstanceComponent(instance) + ";Integrated Security=SSPI;";
}

function Connect(instance) {
  var connection = new ActiveXObject("ADODB.Connection");
  connection.ConnectionString = ConnectionString(instance);
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