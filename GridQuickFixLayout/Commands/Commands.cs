namespace GridQuickFixLayout
{
    [Command(PackageIds.GridColumnIncrementCommand)]
    internal sealed class GridColumnIncrementCommand : BaseCommand<GridColumnIncrementCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            if (documentView.Document.TextBuffer.ContentType.TypeName != ContentTypes.Xaml) return;
            
            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            if (string.IsNullOrWhiteSpace(selectionText)) return;

            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridColumn);
            if (valuesByMatch.Count == 0) return;

            QuickFixLayout.IncrementValues(ref valuesByMatch);

            QuickFixLayout.ReplacePropertyValuesByMatchInSelection(ref selectionText, valuesByMatch, FindableGridProperty.GridColumn);
            QuickFixLayout.ApplySelectionChangesToDocument(selectionText, documentView);
        }
    }

    [Command(PackageIds.GridColumnDecrementCommand)]
    internal sealed class GridColumnDecrementCommand : BaseCommand<GridColumnDecrementCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            if (documentView.Document.TextBuffer.ContentType.TypeName != ContentTypes.Xaml) return;

            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            if (string.IsNullOrWhiteSpace(selectionText)) return;

            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridColumn);
            if (valuesByMatch.Count == 0) return;

            QuickFixLayout.DecrementValues(ref valuesByMatch);

            QuickFixLayout.ReplacePropertyValuesByMatchInSelection(ref selectionText, valuesByMatch, FindableGridProperty.GridColumn);
            QuickFixLayout.ApplySelectionChangesToDocument(selectionText, documentView);
        }
    }

    [Command(PackageIds.GridColumnFillInSequenceCommand)]
    internal sealed class GridColumnFillInSequenceCommand : BaseCommand<GridColumnFillInSequenceCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            if (documentView.Document.TextBuffer.ContentType.TypeName != ContentTypes.Xaml) return;

            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            if (string.IsNullOrWhiteSpace(selectionText)) return;

            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridColumn);
            if (valuesByMatch.Count == 0) return;

            QuickFixLayout.FillInSequence(ref valuesByMatch);

            QuickFixLayout.ReplacePropertyValuesByMatchInSelection(ref selectionText, valuesByMatch, FindableGridProperty.GridColumn);
            QuickFixLayout.ApplySelectionChangesToDocument(selectionText, documentView);
        }
    }

    [Command(PackageIds.GridRowIncrementCommand)]
    internal sealed class GridRowIncrementCommand : BaseCommand<GridRowIncrementCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            if (documentView.Document.TextBuffer.ContentType.TypeName != ContentTypes.Xaml) return;

            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            if (string.IsNullOrWhiteSpace(selectionText)) return;

            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridRow);
            if (valuesByMatch.Count == 0) return;

            QuickFixLayout.IncrementValues(ref valuesByMatch);

            QuickFixLayout.ReplacePropertyValuesByMatchInSelection(ref selectionText, valuesByMatch, FindableGridProperty.GridRow);
            QuickFixLayout.ApplySelectionChangesToDocument(selectionText, documentView);
        }
    }

    [Command(PackageIds.GridRowDecrementCommand)]
    internal sealed class GridRowDecrementCommand : BaseCommand<GridRowDecrementCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            if (documentView.Document.TextBuffer.ContentType.TypeName != ContentTypes.Xaml) return;

            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            if (string.IsNullOrWhiteSpace(selectionText)) return;

            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridRow);
            if (valuesByMatch.Count == 0) return;

            QuickFixLayout.DecrementValues(ref valuesByMatch);

            QuickFixLayout.ReplacePropertyValuesByMatchInSelection(ref selectionText, valuesByMatch, FindableGridProperty.GridRow);
            QuickFixLayout.ApplySelectionChangesToDocument(selectionText, documentView);
        }
    }

    [Command(PackageIds.GridRowFillInSequenceCommand)]
    internal sealed class GridRowFillInSequenceCommand : BaseCommand<GridRowFillInSequenceCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            if (documentView.Document.TextBuffer.ContentType.TypeName != ContentTypes.Xaml) return;

            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            if (string.IsNullOrWhiteSpace(selectionText)) return;

            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridRow);
            if (valuesByMatch.Count == 0) return;

            QuickFixLayout.FillInSequence(ref valuesByMatch);

            QuickFixLayout.ReplacePropertyValuesByMatchInSelection(ref selectionText, valuesByMatch, FindableGridProperty.GridRow);
            QuickFixLayout.ApplySelectionChangesToDocument(selectionText, documentView);
        }
    }
}
