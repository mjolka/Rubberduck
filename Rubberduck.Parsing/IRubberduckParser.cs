using System;
using System.Threading;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser
    {
        RubberduckParserState State { get; }
        void Parse(VBE vbe, CancellationToken cancellationToken);
        void Parse(VBComponent component, CancellationToken cancellationToken);
    }
}