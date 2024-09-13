namespace GridQuickFixLayout
{
    [Command(PackageIds.GridColumnIncrementCommand)]
    internal sealed class GridColumnIncrementCommand : BaseCommand<GridColumnIncrementCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            var documentView = await VS.Documents.GetActiveDocumentViewAsync();
            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridColumn);
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
            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridColumn);
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
            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridColumn);
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
            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridRow);
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
            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridRow);
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
            var selectionText = QuickFixLayout.GetSelectionTextFromDocumentView(documentView);
            var valuesByMatch = QuickFixLayout.GetPropertyValuesByMatchFromSelection(selectionText, FindableGridProperty.GridRow);
            QuickFixLayout.FillInSequence(ref valuesByMatch);

            QuickFixLayout.ReplacePropertyValuesByMatchInSelection(ref selectionText, valuesByMatch, FindableGridProperty.GridRow);
            QuickFixLayout.ApplySelectionChangesToDocument(selectionText, documentView);
        }
    }
}
