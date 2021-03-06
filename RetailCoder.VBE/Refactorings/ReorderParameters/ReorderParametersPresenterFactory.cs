﻿using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenterFactory : IRefactoringPresenterFactory<IReorderParametersPresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IReorderParametersView _view;
        private readonly RubberduckParserState _parseResult;
        private readonly IMessageBox _messageBox;

        public ReorderParametersPresenterFactory(IActiveCodePaneEditor editor, IReorderParametersView view,
            RubberduckParserState parseResult, IMessageBox messageBox)
        {
            _editor = editor;
            _view = view;
            _parseResult = parseResult;
            _messageBox = messageBox;
        }

        public IReorderParametersPresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new ReorderParametersModel(_parseResult, selection.Value, _messageBox);
            return new ReorderParametersPresenter(_view, model, _messageBox);
        }
    }
}
