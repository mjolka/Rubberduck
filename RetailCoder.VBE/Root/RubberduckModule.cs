﻿using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Extensions.Conventions;
using Ninject.Extensions.NamedScope;
using Ninject.Modules;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.Command;
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Root
{
    public class RubberduckModule : NinjectModule
    {
        private readonly IKernel _kernel;
        private readonly VBE _vbe;
        private readonly AddIn _addin;

        public RubberduckModule(IKernel kernel, VBE vbe, AddIn addin)
        {
            _kernel = kernel;
            _vbe = vbe;
            _addin = addin;
        }

        public override void Load()
        {
            Debug.Print("in RubberduckModule.Load()");

            _kernel.Bind<App>().ToSelf().InSingletonScope();

            // bind VBE and AddIn dependencies to host-provided instances.
            _kernel.Bind<VBE>().ToConstant(_vbe);
            _kernel.Bind<AddIn>().ToConstant(_addin);

            BindCodeInspectionTypes();

            var assemblies = new[]
            {
                Assembly.GetExecutingAssembly(),
                Assembly.GetAssembly(typeof(IHostApplication)),
                Assembly.GetAssembly(typeof(IRubberduckParser))
            };

            ApplyConfigurationConvention(assemblies);
            ApplyDefaultInterfacesConvention(assemblies);
            ApplyAbstractFactoryConvention(assemblies);

            Bind<TestExplorerModelBase>().To<StandardModuleTestExplorerModel>().InSingletonScope();

            Bind<IPresenter>().To<TestExplorerDockablePresenter>()
                .WhenInjectedInto<TestExplorerCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new TestExplorerWindow { ViewModel = _kernel.Get<TestExplorerViewModel>() });

            Bind<IPresenter>().To<CodeInspectionsDockablePresenter>()
                .WhenInjectedInto<RunCodeInspectionsCommand>()
                .InSingletonScope()
                .WithConstructorArgument<IDockableUserControl>(new CodeInspectionsWindow { ViewModel = _kernel.Get<InspectionResultsViewModel>() });

            Debug.Print("completed RubberduckModule.Load()");
        }

        private void ApplyDefaultInterfacesConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                // inspections & factories have their own binding rules
                .Where(type => !type.Name.EndsWith("Factory") && !type.GetInterfaces().Contains(typeof(IInspection)))
                .BindDefaultInterface()
                .Configure(binding => binding.InCallScope())); // TransientScope wouldn't dispose disposables
        }

        // note: settings namespace classes are injected in singleton scope
        private void ApplyConfigurationConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllClasses()
                .InNamespaceOf<Configuration>()
                .BindAllInterfaces()
                .Configure(binding => binding.InSingletonScope()));
        }

        // note convention: abstract factory interface names end with "Factory".
        private void ApplyAbstractFactoryConvention(IEnumerable<Assembly> assemblies)
        {
            _kernel.Bind(t => t.From(assemblies)
                .SelectAllInterfaces()
                .Where(type => type.Name.EndsWith("Factory"))
                .BindToFactory()
                .Configure(binding => binding.InSingletonScope()));
        }

        // note: IInspection implementations are discovered in the Rubberduck assembly via reflection.
        private void BindCodeInspectionTypes()
        {
            var inspections = Assembly.GetExecutingAssembly()
                                      .GetTypes()
                                      .Where(type => type.GetInterfaces().Contains(typeof (IInspection)));

            // multibinding for IEnumerable<IInspection> dependency
            foreach (var inspection in inspections)
            {
                _kernel.Bind<IInspection>().To(inspection).InSingletonScope();
            }
        }
    }
}
