﻿using System;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Common;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.ParserErrors;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly IParserErrorsPresenterFactory _parserErrorsPresenterFactory;
        private readonly IRubberduckParser _parser;
        private readonly IInspectorFactory _inspectorFactory;
        private readonly IGeneralConfigService _configService;
        private readonly IAppMenu _appMenus;
        private readonly ParserStateCommandBar _stateBar;
        private readonly KeyHook _hook;

        private readonly Logger _logger;

        private Configuration _config;

        public App(VBE vbe, IMessageBox messageBox,
            IParserErrorsPresenterFactory parserErrorsPresenterFactory,
            IRubberduckParser parser,
            IInspectorFactory inspectorFactory, 
            IGeneralConfigService configService,
            IAppMenu appMenus,
            ParserStateCommandBar stateBar,
            KeyHook hook)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _parserErrorsPresenterFactory = parserErrorsPresenterFactory;
            _parser = parser;
            _inspectorFactory = inspectorFactory;
            _configService = configService;
            _appMenus = appMenus;
            _stateBar = stateBar;
            _hook = hook;
            _logger = LogManager.GetCurrentClassLogger();

            _hook.Attach();
            _hook.KeyPressed += _hook_KeyPressed;
            _configService.SettingsChanged += _configService_SettingsChanged;
        }

        private void _hook_KeyPressed(object sender, KeyHookEventArgs e)
        {
            // We'll add a CancellationToken soon.
            _parser.Parse(e.Component, CancellationToken.None);
        }

        public void Startup()
        {
            CleanReloadConfig();

            _appMenus.Initialize();
            _appMenus.Localize();

            Task.Delay(1000).ContinueWith(t => _parser.Parse(_vbe, CancellationToken.None));
        }

        private void CleanReloadConfig()
        {
            LoadConfig();
            CleanUp();
            Setup();
        }

        private void _configService_SettingsChanged(object sender, EventArgs e)
        {
            CleanReloadConfig();
        }

        private void LoadConfig()
        {
            _logger.Debug("Loading configuration");
            _config = _configService.LoadConfiguration();

            var currentCulture = RubberduckUI.Culture;
            try
            {
                RubberduckUI.Culture = CultureInfo.GetCultureInfo(_config.UserSettings.LanguageSetting.Code);
                _appMenus.Localize();
            }
            catch (CultureNotFoundException exception)
            {
                _logger.Error(exception, "Error Setting Culture for RubberDuck");
                _messageBox.Show(exception.Message, "Rubberduck", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _config.UserSettings.LanguageSetting.Code = currentCulture.Name;
                _configService.SaveConfiguration(_config);
            }
        }

        private void Setup()
        {
            _inspectorFactory.Create();
            _parserErrorsPresenterFactory.Create();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }

            CleanUp();
            _hook.Detach();
        }

        private void CleanUp()
        {
        }
    }
}
