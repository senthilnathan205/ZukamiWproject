Imports Microsoft.Owin
Imports Owin
Imports Microsoft.AspNet.SignalR

<Assembly: OwinStartup(GetType(TMOwinStartup))>

Public Class TMOwinStartup
    Public Sub Configuration(app As IAppBuilder)
        '            GlobalHost.DependencyResolver.Register(GetType(IUserIdProvider), Function() New CustomTMUserIdProvider())
        app.MapSignalR()
    End Sub
End Class



