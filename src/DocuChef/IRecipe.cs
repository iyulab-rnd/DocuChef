﻿using System.Reflection;

namespace DocuChef;

/// <summary>
/// Interface for all document template recipes
/// </summary>
public interface IRecipe : IDisposable
{
    /// <summary>
    /// Adds a variable to the template
    /// </summary>
    void AddVariable(string name, object value);

    /// <summary>
    /// Adds variables from a source object
    /// </summary>
    void AddVariable(object data);

    /// <summary>
    /// Clears all variables from the template
    /// </summary>
    void ClearVariables();

    /// <summary>
    /// Registers a global variable
    /// </summary>
    void RegisterGlobalVariable(string name, object value);
}

/// <summary>
/// Base implementation for document templates
/// </summary>
public abstract class RecipeBase : IRecipe
{
    protected readonly Dictionary<string, object> Variables = new();
    protected readonly Dictionary<string, Func<object>> GlobalVariables = new();
    protected bool IsDisposed;

    /// <summary>
    /// Adds variables from a source object
    /// </summary>
    public virtual void AddVariable(object data)
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data));

        if (data is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                AddVariable(entry.Key.ToString(), entry.Value);
            }
        }
        else
        {
            // Get all properties and fields using extension method
            var properties = data.GetProperties();
            foreach (var kvp in properties)
            {
                AddVariable(kvp.Key, kvp.Value);
            }
        }
    }

    /// <summary>
    /// Adds a variable to the template
    /// </summary>
    public abstract void AddVariable(string name, object value);

    /// <summary>
    /// Clears all variables from the template
    /// </summary>
    public virtual void ClearVariables()
    {
        Variables.Clear();
    }

    /// <summary>
    /// Registers a global variable
    /// </summary>
    public virtual void RegisterGlobalVariable(string name, object value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        if (value is Func<object> valueFactory)
        {
            GlobalVariables[name] = valueFactory;
        }
        else
        {
            GlobalVariables[name] = () => value;
        }
    }

    /// <summary>
    /// Registers standard built-in global variables
    /// </summary>
    protected virtual void RegisterStandardGlobalVariables()
    {
        // Register date/time related variables
        RegisterGlobalVariable("Today", () => DateTime.Today);
        RegisterGlobalVariable("Now", () => DateTime.Now);
        RegisterGlobalVariable("Year", () => DateTime.Now.Year);
        RegisterGlobalVariable("Month", () => DateTime.Now.Month);
        RegisterGlobalVariable("Day", () => DateTime.Now.Day);

        // Register system variables
        RegisterGlobalVariable("MachineName", Environment.MachineName);
        RegisterGlobalVariable("UserName", Environment.UserName);
        RegisterGlobalVariable("OSVersion", Environment.OSVersion.ToString());
        RegisterGlobalVariable("ProcessorCount", Environment.ProcessorCount);
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        IsDisposed = true;
    }

    /// <summary>
    /// Throws an ObjectDisposedException if the object is disposed
    /// </summary>
    protected void ThrowIfDisposed([System.Runtime.CompilerServices.CallerMemberName] string memberName = "")
    {
        if (IsDisposed)
            throw new ObjectDisposedException(GetType().Name, $"Cannot access {memberName} after the object is disposed.");
    }
}