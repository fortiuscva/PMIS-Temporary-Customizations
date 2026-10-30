report 52900 "PMIS Remove Ending Quote"
{
    ApplicationArea = All;
    Caption = 'Remove Ending Quote from Item No.';
    ProcessingOnly = true;
    UsageCategory = Tasks;

    dataset
    {
    }

    trigger OnPreReport()
    begin
        ImportItemNumbers();
        ProcessItemNumbers();
        DownloadProcessingLog();
        ShowSummary();
    end;

    var
        ExcelBuffer: Record "Excel Buffer" temporary;
        ExportExcelBuffer: Record "Excel Buffer" temporary;
        Item: Record Item;
        TempBlob: Codeunit "Temp Blob";
        ItemNumbers: List of [Text];
        ProcessingResults: List of [Text];
        ProcessedItemNumbers: List of [Text];
        FileName: Text;
        SheetName: Text[250];
        DownloadFileName: Text;
        ItemNo: Text[250];
        ItemNoCode: Code[20];
        NewItemNo: Code[20];
        ErrorText: Text;
        InStr: InStream;
        OutStr: OutStream;
        DownloadInStr: InStream;
        UpdatedCount: Integer;
        NotFoundCount: Integer;
        UnchangedCount: Integer;
        DuplicateCount: Integer;
        ConflictCount: Integer;
        InvalidCount: Integer;
        RenameFailedCount: Integer;

        UploadFileMsg: Label 'Select the Excel file that contains the item numbers.';
        NoFileSelectedErr: Label 'No Excel file was selected.';
        InvalidWorkbookErr: Label 'The selected workbook must have No. in the first cell of the first column.';
        EmptyWorkbookErr: Label 'The selected workbook does not contain any item numbers.';
        InvalidItemNoTxt: Label 'Invalid item number: it is longer than 20 characters or blank.';
        DuplicateItemTxt: Label 'Duplicate item number in the Excel file; skipped.';
        ItemNotFoundTxt: Label 'Item not found in Business Central.';
        NoQuoteTxt: Label 'Item number has no ending double quote; unchanged.';
        ConflictTxt: Label 'The cleaned item number already exists in Business Central; skipped.';
        RenameFailedTxt: Label 'Business Central could not rename this item.';
        UpdatedTxt: Label 'Renamed successfully.';
        ResultsSheetNameTxt: Label 'Processing Results', Comment = 'Excel worksheet name.';
        ExcelItemNoColumnTxt: Label 'Excel Item No.', Comment = 'Column heading in the processing results workbook.';
        ResultColumnTxt: Label 'Result', Comment = 'Column heading in the processing results workbook.';
        NewItemNoColumnTxt: Label 'New Item No.', Comment = 'Column heading in the processing results workbook.';
        ResultsReportHeaderTxt: Label 'Item Number Processing', Comment = 'Report header in the processing results workbook.';
        ResultsFileNameTxt: Label 'Item No. Processing Results.xlsx';
        SummaryMsg: Label 'Item number processing is complete.\Renamed: %1\Not found: %2\Already clean: %3\Duplicate Excel rows: %4\Rename conflicts: %5\Invalid rows: %6\Rename failures: %7\A detailed result log was downloaded.';

    local procedure ImportItemNumbers()
    begin
        if not UploadIntoStream(UploadFileMsg, '', '', FileName, InStr) then
            Error(NoFileSelectedErr);

        TempBlob.CreateOutStream(OutStr);
        CopyStream(OutStr, InStr);

        TempBlob.CreateInStream(InStr);
        SheetName := ExcelBuffer.SelectSheetsNameStream(InStr);
        TempBlob.CreateInStream(InStr);
        ErrorText := ExcelBuffer.OpenBookStream(InStr, SheetName);
        if ErrorText <> '' then
            Error(ErrorText);

        ExcelBuffer.ReadSheet();
        ExcelBuffer.CloseBook();

        ExcelBuffer.SetRange("Row No.", 1);
        ExcelBuffer.SetRange("Column No.", 1);
        if not ExcelBuffer.FindFirst() then
            Error(InvalidWorkbookErr);
        if LowerCase(DelChr(ExcelBuffer."Cell Value as Text", '<>', ' ')) <> 'no.' then
            Error(InvalidWorkbookErr);

        ExcelBuffer.Reset();
        ExcelBuffer.SetRange("Column No.", 1);
        ExcelBuffer.SetFilter("Row No.", '>1');
        if ExcelBuffer.FindSet() then
            repeat
                ItemNo := DelChr(ExcelBuffer."Cell Value as Text", '<>', ' ');
                if ItemNo <> '' then
                    ItemNumbers.Add(ItemNo);
            until ExcelBuffer.Next() = 0;

        if ItemNumbers.Count() = 0 then
            Error(EmptyWorkbookErr);
    end;

    local procedure ProcessItemNumbers()
    var
        Index: Integer;
    begin
        foreach ItemNo in ItemNumbers do begin
            if (ItemNo = '') or (StrLen(ItemNo) > MaxStrLen(ItemNoCode)) then begin
                ProcessingResults.Add(InvalidItemNoTxt);
                InvalidCount += 1;
            end else begin
                ItemNoCode := ItemNo;
                if ProcessedItemNumbers.Contains(ItemNoCode) then begin
                    ProcessingResults.Add(DuplicateItemTxt);
                    DuplicateCount += 1;
                end else begin
                    ProcessedItemNumbers.Add(ItemNoCode);
                    ProcessItemNumber(ItemNoCode);
                end;
            end;
        end;
    end;

    local procedure ProcessItemNumber(CurrentItemNo: Code[20])
    begin
        if not Item.Get(CurrentItemNo) then begin
            ProcessingResults.Add(ItemNotFoundTxt);
            NotFoundCount += 1;
            exit;
        end;

        if CopyStr(CurrentItemNo, StrLen(CurrentItemNo), 1) <> '"' then begin
            ProcessingResults.Add(NoQuoteTxt);
            UnchangedCount += 1;
            exit;
        end;

        NewItemNo := CopyStr(CurrentItemNo, 1, StrLen(CurrentItemNo) - 1);
        if NewItemNo = '' then begin
            ProcessingResults.Add(InvalidItemNoTxt);
            InvalidCount += 1;
            exit;
        end;

        if Item.Get(NewItemNo) then begin
            ProcessingResults.Add(ConflictTxt);
            ConflictCount += 1;
            exit;
        end;

        if Item.Rename(NewItemNo) then begin
            ProcessingResults.Add(UpdatedTxt);
            UpdatedCount += 1;
        end else begin
            ProcessingResults.Add(RenameFailedTxt);
            RenameFailedCount += 1;
        end;
    end;

    local procedure DownloadProcessingLog()
    var
        Index: Integer;
    begin
        ExportExcelBuffer.CreateNewBook(ResultsSheetNameTxt);
        ExportExcelBuffer.NewRow();
        AddLogColumn(ExcelItemNoColumnTxt, true);
        AddLogColumn(ResultColumnTxt, true);
        AddLogColumn(NewItemNoColumnTxt, true);

        for Index := 1 to ItemNumbers.Count() do begin
            ExportExcelBuffer.NewRow();
            AddLogColumn(ItemNumbers.Get(Index), false);
            AddLogColumn(ProcessingResults.Get(Index), false);
            AddLogColumn(GetNewItemNo(ItemNumbers.Get(Index), ProcessingResults.Get(Index)), false);
        end;

        ExportExcelBuffer.WriteSheet(ResultsReportHeaderTxt, CompanyName(), UserId());
        ExportExcelBuffer.CloseBook();
        TempBlob.CreateOutStream(OutStr);
        ExportExcelBuffer.SaveToStream(OutStr, true);
        TempBlob.CreateInStream(DownloadInStr);
        DownloadFileName := ResultsFileNameTxt;
        DownloadFromStream(DownloadInStr, '', '', '', DownloadFileName);
    end;

    local procedure AddLogColumn(Value: Text; IsBold: Boolean)
    begin
        ExportExcelBuffer.AddColumn(Value, false, '', IsBold, false, false, '', ExportExcelBuffer."Cell Type"::Text);
    end;

    local procedure GetNewItemNo(CurrentItemNo: Text; ProcessingResult: Text): Text
    begin
        if ProcessingResult <> UpdatedTxt then
            exit('');

        exit(CopyStr(CurrentItemNo, 1, StrLen(CurrentItemNo) - 1));
    end;

    local procedure ShowSummary()
    begin
        Message(
            SummaryMsg,
            UpdatedCount,
            NotFoundCount,
            UnchangedCount,
            DuplicateCount,
            ConflictCount,
            InvalidCount,
            RenameFailedCount);
    end;
}
