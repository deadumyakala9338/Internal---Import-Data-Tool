pageextension 50148 "VAT Entry TJP" extends "VAT Entries"
{
    actions
    {
        addafter("&Navigate_Promoted")
        {
            actionref(VATEntryImport_Promoted; VATEntryImport) { }
        }
        addafter("&Navigate")
        {
            action(VATEntryImport)
            {
                ApplicationArea = All;
                Caption = 'VAT Entry Data Import';
                trigger OnAction()
                var
                    VATEntryImportTJP: Codeunit "VAT Entries Import TJP";
                begin
                    VATEntryImportTJP.ImportVATEntryExtensionFields();
                end;
            }
        }
    }
}