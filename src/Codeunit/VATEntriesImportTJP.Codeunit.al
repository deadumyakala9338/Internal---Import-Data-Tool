codeunit 50148 "VAT Entries Import TJP"
{
    Permissions = tabledata "VAT Entry" = RM;

    procedure ImportVATEntryExtensionFields()
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        VATEntry: Record "VAT Entry";
        InStream: InStream;
        Filename: Text;
        RowNo: Integer;
        ColumnNo: Integer;
        LastRowNo: Integer;
        EntryNo: Integer;
        NoTypeTJP: Enum "No. Type VAT Entry TJP";
        BillPayToNoExtendedTJP: Code[20];
        NoOfRecrodsUpdated: Integer;
        SheetNameTxt: Label 'VAT Entry', Locked = true;
        SelectFileMsg: Label 'Select the excel file...';
        NoOfRecrdsUpdatedMsg: Label 'Number of records modified: %1', Comment = '%1, number of records';
    begin
        if UploadIntoStream(SelectFileMsg, '', '', Filename, InStream) then begin
            TempExcelBuffer.OpenBookStream(InStream, SheetNameTxt);
            TempExcelBuffer.ReadSheet();
            TempExcelBuffer.SetRange("Column No.", 1);
            TempExcelBuffer.FindLast();
            LastRowNo := TempExcelBuffer."Row No.";
            TempExcelBuffer.Reset();
            for RowNo := 4 to LastRowNo do begin
                ColumnNo := 1;
                for ColumnNo := 1 to 3
                do
                    if TempExcelBuffer.Get(RowNo, ColumnNo) then
                        if TempExcelBuffer."Cell Value as Text" <> '' then
                            case ColumnNo of
                                1:
                                    Evaluate(EntryNo, TempExcelBuffer."Cell Value as Text");
                                2:
                                    begin
                                        Evaluate(NoTypeTJP, TempExcelBuffer."Cell Value as Text");
                                        if VATEntry.Get(EntryNo) = true then begin
                                            VATEntry."No. Type TJP" := NoTypeTJP;
                                            if VATEntry.Modify(false) = true then
                                                NoOfRecrodsUpdated := NoOfRecrodsUpdated + 1;
                                        end;
                                    end;
                            /* 3:
                                begin
                                    Evaluate(BillPayToNoExtendedTJP, TempExcelBuffer."Cell Value as Text");
                                    if VATEntry.Get(EntryNo) = true then begin
                                        VATEntry."No. Type TJP" := NoTypeTJP;
                                        VATEntry."Bill/Pay-to No. (Extended) TJP" := BillPayToNoExtendedTJP;
                                        if VATEntry.Modify(false) = true then
                                            NoOfRecrodsUpdated := NoOfRecrodsUpdated + 1;
                                    end;
                                end; */
                            end
                        else
                            DoNothing()
                    else
                        DoNothing();
            end;
        end;
        Message(NoOfRecrdsUpdatedMsg, NoOfRecrodsUpdated);
    end;

    local procedure DoNothing()
    var
    begin
    end;
}