report 50140 "Import/Update Ship-to Phone"
{
    Caption = 'Import/Update Ship-to Phone';
    ProcessingOnly = true;
    UsageCategory = Administration;
    ApplicationArea = All;

    dataset
    {
        dataitem(ImportData; Integer)
        {
            DataItemTableView = sorting(Number);

            trigger OnPreDataItem()
            var
                UpdSalesShiptoPhone: Codeunit "Upd. Sales Ship-to Phone";
            begin
                if ImportType = ImportType::"Sales Header" then
                    UpdSalesShiptoPhone.UpdateSalesHeaderShipToPhone();
                if ImportType = ImportType::"Sales Shipment Header" then
                    UpdSalesShiptoPhone.UpdateSalesShipHeaderShipToPhone();
                if ImportType = ImportType::"Sales Invoice Header" then
                    UpdSalesShiptoPhone.UpdateSalesInvHeaderShipToPhone();
                if ImportType = ImportType::"Sales Cr. Memo Header" then
                    UpdSalesShiptoPhone.UpdateSalesCrMemoHeaderShipToPhone();
            end;
        }
    }

    requestpage
    {
        SaveValues = true;

        layout
        {
            area(content)
            {
                group(Options)
                {
                    Caption = 'Options';
                    field(ReqImportType; ImportType)
                    {
                        ApplicationArea = All;
                        Caption = 'Import Type';
                    }
                }
            }
        }
    }

    trigger OnPostReport()
    begin
    end;

    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        FileManagement: Codeunit "File Management";
        TempBlob: Codeunit "Temp Blob";
        FileInStream: InStream;
        FileExt: Text;
        FileName: Text;
        SheetName: Text[250];
        ImportType: Option " ","Sales Header","Sales Shipment Header","Sales Invoice Header","Sales Cr. Memo Header";
        DialogTxt: Label 'Import (%1)|%1';
        FilterTxt: Label '*.xlsx;*.xls;*.*', Locked = true;
        Text001Lbl: Label 'File from excel';
        Text004Lbl: Label 'You must enter a file name.';
        Text007Lbl: Label 'You must enter an excel worksheet name.';
}

