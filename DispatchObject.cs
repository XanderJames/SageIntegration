using System;
using System.Reflection;

/*
 * This object was given to me by someone on the sage forums. 
 * object wraps the unsafe bindings using reflection
 * reflection is a bit slower, but it seems to work ok.
 */
public class DispatchObject : IDisposable
{
    protected object m_object = null;

    private BindingFlags m_flgMethod = BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase | BindingFlags.InvokeMethod;
    private BindingFlags m_flgGetProperty = BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase | BindingFlags.GetProperty;
    private BindingFlags m_flgSetProperty = BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase | BindingFlags.SetProperty;

    protected DispatchObject()
    {
    }

    public DispatchObject(string sProgId)
    {
        m_object = Activator.CreateInstance(Type.GetTypeFromProgID(sProgId, true));
    }

    public DispatchObject(string sProgId, string sServer)
    {
        m_object = Activator.CreateInstance(Type.GetTypeFromProgID(sProgId, sServer, true));
    }

    public DispatchObject(Guid guid)
    {
        m_object = Activator.CreateInstance(Type.GetTypeFromCLSID(guid, true));
    }

    public DispatchObject(Guid guid, string sServer)
    {
        m_object = Activator.CreateInstance(Type.GetTypeFromCLSID(guid, sServer, true));
    }

    public DispatchObject(object o)
    {
        m_object = o;
    }

    ~DispatchObject()
    {
        Dispose();
    }

    public object InvokeMethodByRef(string sMethodName, object[] aryParams)
    {
        ParameterModifier[] pmods = new ParameterModifier[] { GetParameterModifier(aryParams.Length, true) };
        return m_object.GetType().InvokeMember(sMethodName, m_flgMethod, null, m_object, aryParams, pmods, null, null);
    }

    public object InvokeMethod(string sMethodName, params object[] aryParams)
    {
        return m_object.GetType().InvokeMember(sMethodName, m_flgMethod, null, m_object, aryParams);
    }

    public object InvokeMethod(string sMethodName, object[] aryParams, ParameterModifier[] pmods)
    {
        return m_object.GetType().InvokeMember(sMethodName, m_flgMethod, null, m_object, aryParams, pmods, null, null);
    }

    public object GetProperty(string sPropertyName)
    {
        return m_object.GetType().InvokeMember(sPropertyName, m_flgGetProperty, null, m_object, null);
    }

    public object SetProperty(string sPropertyName, object oValue)
    {
        return m_object.GetType().InvokeMember(sPropertyName, m_flgSetProperty, null, m_object, new object[] { oValue });
    }

    public object GetObject()
    {
        return m_object;
    }

    virtual public void Dispose()
    {
        if (m_object != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_object);
            m_object = null;
            GC.SuppressFinalize(this);
        }
    }

    public static ParameterModifier GetParameterModifier(int parameterCount, bool initialValue)
    {
        ParameterModifier mod = new ParameterModifier(parameterCount);
        for (int x = 0; x < parameterCount; x++) mod[x] = initialValue;
        return mod;
    }
}
