﻿using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.ParserProgress
{
    public interface IParsingProgressPresenter
    {
        IEnumerable<Declaration> Parse(VBProject project);
    }

    public class ParsingProgressPresenter : IParsingProgressPresenter
    {
        private readonly IRubberduckParser _parser;

        public ParsingProgressPresenter(IRubberduckParser parser)
        {
            _parser = parser;
        }

        public IEnumerable<Declaration> Parse(VBProject project)
        {
            using (var view = new ProgressDialog(new ParserProgessViewModel(_parser, project)))
            {
                view.ShowDialog();
                return view.Result.AllDeclarations;
            }
        }
    }
}
