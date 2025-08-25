codeunit 50149 "Upd. Sales Ship-to Phone"
{
    Permissions = tabledata "Sales Shipment Header" = RM,
                  tabledata "Sales Invoice Header" = RM,
                  tabledata "Sales Cr.Memo Header" = RM;
    procedure UpdateSalesHeaderShipToPhone()
    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        SalesHeader: Record "Sales Header";
        InStream: InStream;
        OrderTypeTxt: Text[30];
        OrderType: Enum "Sales Document Type";
        OrderNo: Code[20];
        ShipToPhoneNo: Code[30];
        Filename: Text;
        LastRowNo: Integer;
        RowNo: Integer;
        ColumnNo: Integer;
        UpdRecordCnt: Integer;
        SheetNameLbl: Label 'Sales Header', Locked = true; // Sheet name in the excel file should be 'Sales Header'
        SelectFileMsg: Label 'Select the excel file...';
        UpdateRecMsg: Label 'Number of records modified: %1', Comment = '%1, number of records';
    begin
        if UploadIntoStream(SelectFileMsg, '', '', Filename, InStream) then begin
            TempExcelBuf.OpenBookStream(InStream, SheetNameLbl);
            TempExcelBuf.ReadSheet();
            TempExcelBuf.SetRange("Column No.", 1);
            TempExcelBuf.FindLast();
            LastRowNo := TempExcelBuf."Row No.";
            TempExcelBuf.Reset();
            for RowNo := 4 to LastRowNo do begin
                ColumnNo := 1;
                for ColumnNo := 1 to 3 do
                    if TempExcelBuf.Get(RowNo, ColumnNo) then
                        if TempExcelBuf."Cell Value as Text" <> '' then
                            case ColumnNo of
                                1:
                                    Evaluate(OrderTypeTxt, TempExcelBuf."Cell Value as Text");
                                2:
                                    Evaluate(OrderNo, TempExcelBuf."Cell Value as Text");
                                3:
                                    begin
                                        Evaluate(ShipToPhoneNo, TempExcelBuf."Cell Value as Text");
                                        OrderType := TextToEnumConversion(OrderTypeTxt);
                                        SalesHeader.Reset();
                                        SalesHeader.SetRange("Document Type", OrderType);
                                        SalesHeader.SetRange("No.", OrderNo);
                                        if SalesHeader.FindSet() then begin
                                            SalesHeader."Ship-to Phone No." := ShipToPhoneNo;
                                            if SalesHeader.Modify() then
                                                UpdRecordCnt := UpdRecordCnt + 1;
                                        end;
                                    end;
                            end;
            end;
            Message(UpdateRecMsg, UpdRecordCnt);
        end;
    end;

    /* procedure UpdSalesHeaderShipToPhone()
    var
        SalesHeader: Record "Sales Header";
        OrderType: Enum "Sales Document Type";
        UpdRecCnt: Integer;
        Progress: Dialog;
        UpdateRecMsg: Label 'Number of records modified: %1', Comment = '%1, number of records';
        ProgLbl: Label 'Processing....#1###################\';
    begin
        Progress.Open(ProgLbl);
        SalesHeader.Reset();
        SalesHeader.SetCurrentKey("Document Type", "No.");
        SalesHeader.SetRange("Ship-to Phone No. MJP", '<>', '');
        if SalesHeader.FindSet() then begin
            repeat
                UpdRecCnt += 1;
                Progress.Update(1, UpdRecCnt);
                SalesHeader.Validate("Ship-to Phone No.", SalesHeader."Ship-to Phone No. MJP");
                SalesHeader.Modify();
            until SalesHeader.Next() = 0;
            Message(UpdateRecMsg, UpdRecCnt);
        end;
        Progress.Close();
    end; */

    procedure UpdateSalesShipHeaderShipToPhone()
    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        SalesShipHeader: Record "Sales Shipment Header";
        InStream: InStream;
        PstdShipmtNo: Code[20];
        ShipToPhoneNo: Code[30];
        Filename: Text;
        LastRowNo: Integer;
        RowNo: Integer;
        ColumnNo: Integer;
        UpdRecordCnt: Integer;
        SheetNameLbl: Label 'Sales Shipment Header', Locked = true; // Sheet name in the excel file should be 'Sales Shipment Header'
        SelectFileMsg: Label 'Select the excel file...';
        UpdateRecMsg: Label 'Number of records modified: %1', Comment = '%1, number of records';
    begin
        if UploadIntoStream(SelectFileMsg, '', '', Filename, InStream) then begin
            TempExcelBuf.OpenBookStream(InStream, SheetNameLbl);
            TempExcelBuf.ReadSheet();
            TempExcelBuf.SetRange("Column No.", 1);
            TempExcelBuf.FindLast();
            LastRowNo := TempExcelBuf."Row No.";
            TempExcelBuf.Reset();
            for RowNo := 4 to LastRowNo do begin
                ColumnNo := 1;
                for ColumnNo := 1 to 2
                do
                    if TempExcelBuf.Get(RowNo, ColumnNo) then
                        if TempExcelBuf."Cell Value as Text" <> '' then
                            case ColumnNo of
                                1:
                                    Evaluate(PstdShipmtNo, TempExcelBuf."Cell Value as Text");
                                2:
                                    begin
                                        Evaluate(ShipToPhoneNo, TempExcelBuf."Cell Value as Text");
                                        if SalesShipHeader.Get(PstdShipmtNo) then begin
                                            SalesShipHeader."Ship-to Phone No." := ShipToPhoneNo;
                                            if SalesShipHeader.Modify() then
                                                UpdRecordCnt := UpdRecordCnt + 1;
                                        end;
                                    end;
                            end
                        else
                            DoNothing()
                    else
                        DoNothing();
            end;
            Message(UpdateRecMsg, UpdRecordCnt);
        end;
    end;

    procedure UpdateSalesInvHeaderShipToPhone()
    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        SalesInvHeader: Record "Sales Invoice Header";
        InStream: InStream;
        PstdInvoiceNo: Code[20];
        ShipToPhoneNo: Code[30];
        Filename: Text;
        LastRowNo: Integer;
        RowNo: Integer;
        ColumnNo: Integer;
        UpdRecordCnt: Integer;
        SheetNameLbl: Label 'Sales Invoice Header', Locked = true; // Sheet name in the excel file should be 'Sales Invoice Header'
        SelectFileMsg: Label 'Select the excel file...';
        UpdateRecMsg: Label 'Number of records modified: %1', Comment = '%1, number of records';
    begin
        if UploadIntoStream(SelectFileMsg, '', '', Filename, InStream) then begin
            TempExcelBuf.OpenBookStream(InStream, SheetNameLbl);
            TempExcelBuf.ReadSheet();
            TempExcelBuf.SetRange("Column No.", 1);
            TempExcelBuf.FindLast();
            LastRowNo := TempExcelBuf."Row No.";
            TempExcelBuf.Reset();
            for RowNo := 4 to LastRowNo do begin
                ColumnNo := 1;
                for ColumnNo := 1 to 2
                do
                    if TempExcelBuf.Get(RowNo, ColumnNo) then
                        if TempExcelBuf."Cell Value as Text" <> '' then
                            case ColumnNo of
                                1:
                                    Evaluate(PstdInvoiceNo, TempExcelBuf."Cell Value as Text");
                                2:
                                    begin
                                        Evaluate(ShipToPhoneNo, TempExcelBuf."Cell Value as Text");
                                        if SalesInvHeader.Get(PstdInvoiceNo) then begin
                                            SalesInvHeader."Ship-to Phone No." := ShipToPhoneNo;
                                            if SalesInvHeader.Modify() then
                                                UpdRecordCnt := UpdRecordCnt + 1;
                                        end;
                                    end;
                            end
                        else
                            DoNothing()
                    else
                        DoNothing();
            end;
            Message(UpdateRecMsg, UpdRecordCnt);
        end;
    end;

    procedure UpdateSalesCrMemoHeaderShipToPhone()
    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        SalesCrMemoHeader: Record "Sales Cr.Memo Header";
        InStream: InStream;
        PstdCrMemoNo: Code[20];
        ShipToPhoneNo: Code[30];
        Filename: Text;
        LastRowNo: Integer;
        RowNo: Integer;
        ColumnNo: Integer;
        UpdRecordCnt: Integer;
        SheetNameLbl: Label 'Sales Cr.Memo Header', Locked = true; // Sheet name in the excel file should be 'Sales Cr.Memo Header'
        SelectFileMsg: Label 'Select the excel file...';
        UpdateRecMsg: Label 'Number of records modified: %1', Comment = '%1, number of records';
    begin
        if UploadIntoStream(SelectFileMsg, '', '', Filename, InStream) then begin
            TempExcelBuf.OpenBookStream(InStream, SheetNameLbl);
            TempExcelBuf.ReadSheet();
            TempExcelBuf.SetRange("Column No.", 1);
            TempExcelBuf.FindLast();
            LastRowNo := TempExcelBuf."Row No.";
            TempExcelBuf.Reset();
            for RowNo := 4 to LastRowNo do begin
                ColumnNo := 1;
                for ColumnNo := 1 to 2
                do
                    if TempExcelBuf.Get(RowNo, ColumnNo) then
                        if TempExcelBuf."Cell Value as Text" <> '' then
                            case ColumnNo of
                                1:
                                    Evaluate(PstdCrMemoNo, TempExcelBuf."Cell Value as Text");
                                2:
                                    begin
                                        Evaluate(ShipToPhoneNo, TempExcelBuf."Cell Value as Text");
                                        if SalesCrMemoHeader.Get(PstdCrMemoNo) then begin
                                            SalesCrMemoHeader."Ship-to Phone No." := ShipToPhoneNo;
                                            if SalesCrMemoHeader.Modify() then
                                                UpdRecordCnt := UpdRecordCnt + 1;
                                        end;
                                    end;
                            end
                        else
                            DoNothing()
                    else
                        DoNothing();
            end;
            Message(UpdateRecMsg, UpdRecordCnt);
        end;
    end;

    local procedure DoNothing()
    var
    begin
    end;

    procedure TextToEnumConversion(TextValue: Text): Enum "Sales Document Type"
    var
        OrderStatus: Enum "Sales Document Type";
    begin

        case TextValue of
            'Quote':
                OrderStatus := OrderStatus::Quote;
            'Order':
                OrderStatus := OrderStatus::Order;
            'Invoice':
                OrderStatus := OrderStatus::Invoice;
            'Credit Memo':
                OrderStatus := OrderStatus::"Credit Memo";
            'Blanket Order':
                OrderStatus := OrderStatus::"Blanket Order";
            'Return Order':
                OrderStatus := OrderStatus::"Return Order";
            else
                Error('Invalid Text Value: %1', TextValue);
        end;
        exit(OrderStatus);
    end;

}