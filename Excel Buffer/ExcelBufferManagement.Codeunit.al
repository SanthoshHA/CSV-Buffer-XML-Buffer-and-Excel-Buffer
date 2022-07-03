codeunit 50102 "Excel Buffer Management"
{
    trigger OnRun()
    begin
        //ExportItemsToExcelUsingExcelBuffer();
        ImportItemsFromExcelUsingExcelBuffer();
    end;

    local procedure ExportItemsToExcelUsingExcelBuffer()
    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        ExcelFileNameLbl: Label 'Items_%1_%2';
    begin
        TempExcelBuf.DeleteAll();
        FillExcelHeader(TempExcelBuf);
        FillExcelBuffer(TempExcelBuf);

        TempExcelBuf.CreateNewBook('Items');
        TempExcelBuf.WriteSheet('Items', CompanyName(), UserId());
        TempExcelBuf.CloseBook();
        TempExcelBuf.SetFriendlyFilename(StrSubstNo(ExcelFileNameLbl, CurrentDateTime, UserId));
        TempExcelBuf.OpenExcel();
    end;

    local procedure FillExcelHeader(var TempExcelBuf: Record "Excel Buffer" temporary)
    begin
        TempExcelBuf.NewRow();
        TempExcelBuf.AddColumn('No', false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
        TempExcelBuf.AddColumn('Description', false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
        TempExcelBuf.AddColumn('Inventory', false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
    end;

    local procedure FillExcelBuffer(var TempExcelBuf: Record "Excel Buffer" temporary)
    var
        Item: Record Item;
    begin
        Item.SetAutoCalcFields(Inventory);
        if Item.FindSet() then
            repeat
                TempExcelBuf.NewRow();
                TempExcelBuf.AddColumn(Item."No.", false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
                TempExcelBuf.AddColumn(Item.Description, false, '', false, false, false, '', TempExcelBuf."Cell Type"::Text);
                TempExcelBuf.AddColumn(Item.Inventory, false, '', false, false, false, '', TempExcelBuf."Cell Type"::Number);
            until Item.Next() = 0;
    end;

    local procedure ImportItemsFromExcelUsingExcelBuffer()
    var
        Item: Record Item;
        Filename: Text;
        Sheetname: Text;
        RowNo: Integer;
        TotalRows: Integer;
        Instr: InStream;
    begin
        ExcelBuf.DeleteAll();
        if not UploadIntoStream('Select File to Upload', '', '', Filename, Instr) then
            exit;

        if Filename <> '' then
            Sheetname := ExcelBuf.SelectSheetsNameStream(Instr)
        else
            exit;

        ExcelBuf.Reset;
        ExcelBuf.OpenBookStream(Instr, Sheetname);
        ExcelBuf.ReadSheet();

        TotalRows := GetTotalCount();

        // Reading from 2nd row (First row was field caption)
        for RowNo := 2 to TotalRows do begin
            if not Item.Get(GetValueAtIndex(RowNo, 1)) then begin
                Item.Init();
                Item.Validate("No.", GetValueAtIndex(RowNo, 1));
                Item.Insert(true);
            end;
            Item.Validate(Description, GetValueAtIndex(RowNo, 2));
            Item.Modify(true);
        end;

        Message('%1 Rows Imported Successfully!!', TotalRows - 1);
    end;

    local procedure GetValueAtIndex(RowNo: Integer; ColNo: Integer): Text
    begin
        IF ExcelBuf.Get(RowNo, ColNo) then
            exit(ExcelBuf."Cell Value as Text");
    end;

    local procedure GetTotalCount(): Integer
    begin
        ExcelBuf.Reset();
        ExcelBuf.SetRange("Column No.", 1);
        Exit(ExcelBuf.Count);
    end;

    var
        ExcelBuf: Record "Excel Buffer";
}