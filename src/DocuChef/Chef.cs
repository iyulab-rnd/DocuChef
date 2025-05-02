using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Reflection;
using DocuChef.Excel;
using DocuChef.PowerPoint;
using DocuChef.Word;
using DocuChef.Utils;

namespace DocuChef
{
    /// <summary>
    /// Main entry point for the DocuChef library - creates template engines for different document types
    /// </summary>
    public class Chef
    {
        private readonly Dictionary<string, object> _globalSettings;
        private readonly LogCallback? _logCallback;

        /// <summary>
        /// Creates a new instance of DocuChef
        /// </summary>
        public Chef()
        {
            _globalSettings = new Dictionary<string, object>();
        }

        /// <summary>
        /// Creates a new instance of DocuChef with logging callback
        /// </summary>
        public Chef(LogCallback logCallback) : this()
        {
            _logCallback = logCallback;
            LoggingHelper.SetLogCallback(logCallback);
        }

        /// <summary>
        /// Creates a new instance of DocuChef with global settings
        /// </summary>
        public Chef(Dictionary<string, object> globalSettings)
        {
            _globalSettings = globalSettings ?? new Dictionary<string, object>();
        }

        /// <summary>
        /// Creates a new instance of DocuChef with global settings and logging callback
        /// </summary>
        public Chef(Dictionary<string, object> globalSettings, LogCallback logCallback)
        {
            _globalSettings = globalSettings ?? new Dictionary<string, object>();
            _logCallback = logCallback;
            LoggingHelper.SetLogCallback(logCallback);
        }

        /// <summary>
        /// Loads a document template based on its file extension
        /// </summary>
        public IRecipe LoadRecipe(string templatePath, RecipeOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(templatePath);

            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            var extension = Path.GetExtension(templatePath).ToLowerInvariant();

            return extension switch
            {
                ".xlsx" => LoadExcelRecipe(templatePath, options),
                ".docx" => LoadWordRecipe(templatePath, options),
                ".pptx" => LoadPowerPointRecipe(templatePath, options),
                _ => throw new NotSupportedException($"Unsupported file format: {extension}")
            };
        }

        /// <summary>
        /// Loads an Excel template file
        /// </summary>
        public ExcelRecipe LoadExcelRecipe(string templatePath, RecipeOptions? options = null)
        {
            var finalOptions = PrepareOptions(options);
            LoggingHelper.LogInformation($"Loading Excel recipe: {templatePath}");
            return new ExcelRecipe(templatePath, finalOptions);
        }

        /// <summary>
        /// Loads a Word template file
        /// </summary>
        public WordRecipe LoadWordRecipe(string templatePath, RecipeOptions? options = null)
        {
            var finalOptions = PrepareOptions(options);
            LoggingHelper.LogInformation($"Loading Word recipe: {templatePath}");
            return new WordRecipe(templatePath, finalOptions);
        }

        /// <summary>
        /// Loads a PowerPoint template file
        /// </summary>
        public PowerPointRecipe LoadPowerPointRecipe(string templatePath, RecipeOptions? options = null)
        {
            var finalOptions = PrepareOptions(options);
            LoggingHelper.LogInformation($"Loading PowerPoint recipe: {templatePath}");
            return new PowerPointRecipe(templatePath, finalOptions);
        }

        /// <summary>
        /// Adds or updates a global setting
        /// </summary>
        /// <param name="key">Setting key</param>
        /// <param name="value">Setting value</param>
        /// <returns>Current DocuChef instance for chaining</returns>
        public Chef AddGlobalSetting(string key, object value)
        {
            _globalSettings[key] = value;
            LoggingHelper.LogInformation($"Added global setting: {key}");
            return this;
        }

        /// <summary>
        /// Gets a global setting
        /// </summary>
        /// <param name="key">Setting key</param>
        /// <returns>Setting value or null if not found</returns>
        public object? GetGlobalSetting(string key)
        {
            return _globalSettings.TryGetValue(key, out var value) ? value : null;
        }

        /// <summary>
        /// Sets the logging callback for this DocuChef instance
        /// </summary>
        /// <param name="logCallback">New logging callback</param>
        /// <returns>Current DocuChef instance for chaining</returns>
        public Chef SetLogCallback(LogCallback logCallback)
        {
            LoggingHelper.SetLogCallback(logCallback);
            return this;
        }

        /// <summary>
        /// Prepares options and applies global settings and logging callback
        /// </summary>
        /// <param name="options">Original options (can be null)</param>
        /// <returns>Prepared options</returns>
        private RecipeOptions PrepareOptions(RecipeOptions? options)
        {
            var finalOptions = options ?? new RecipeOptions();

            // Apply logging callback
            if (_logCallback != null && finalOptions.LogCallback == null)
            {
                finalOptions.LogCallback = _logCallback;
                LoggingHelper.LogInformation("Applied global logging callback to options");
            }

            // Apply global settings
            ApplyGlobalSettings(finalOptions);

            return finalOptions;
        }

        /// <summary>
        /// Applies global settings to options
        /// </summary>
        /// <param name="options">Options to apply settings to</param>
        private void ApplyGlobalSettings(RecipeOptions options)
        {
            // Apply key global settings to recipe options
            foreach (var setting in _globalSettings)
            {
                LoggingHelper.LogInformation($"Applying global setting: {setting.Key}");

                switch (setting.Key)
                {
                    case "DefaultCulture" when setting.Value is CultureInfo cultureInfo:
                        options.CultureInfo = cultureInfo;
                        break;
                    case "DefaultNullDisplay" when setting.Value is string nullDisplayStr:
                        options.NullDisplayString = nullDisplayStr;
                        break;
                    default:
                        // Try to set property by reflection if it exists
                        var property = options.GetType().GetProperty(setting.Key);
                        if (property != null && property.CanWrite &&
                            property.PropertyType.IsAssignableFrom(setting.Value.GetType()))
                        {
                            property.SetValue(options, setting.Value);
                            LoggingHelper.LogInformation($"Set {setting.Key} from global settings");
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// Factory method: Creates a default DocuChef instance
        /// </summary>
        /// <returns>A new DocuChef instance</returns>
        public static Chef Create()
        {
            return new Chef();
        }

        /// <summary>
        /// Factory method: Creates a DocuChef instance with logging enabled
        /// </summary>
        /// <param name="logCallback">Logging callback</param>
        /// <returns>A new DocuChef instance</returns>
        public static Chef CreateWithLogging(LogCallback logCallback)
        {
            return new Chef(logCallback);
        }
    }
}