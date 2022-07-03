codeunit 50101 "CSV Buffer Management"
{
    trigger OnRun()
    begin
        //ExportItemsToCSVUsingCSVBuffer();
        //ImportItemsFromCSVUsingCSVBuffer(); //Simple
        ImportSalesOrderDocumentsFromTextUsingCSVBuffer(); //Complex
    end;

    local procedure ExportItemsToCSVUsingCSVBuffer()
    var
        Item: Record Item;
        TempCSVBuffer: Record "CSV Buffer" temporary;
        TempBlob: Codeunit "Temp Blob";
        InStreamL: InStream;
        FileName: Text;
        LineNo: Integer;
    begin
        Item.SetAutoCalcFields(Inventory);
        if Item.FindSet() then
            repeat
                LineNo += 1;
                TempCSVBuffer.InsertEntry(LineNo, 1, Item."No.");
                TempCSVBuffer.InsertEntry(LineNo, 2, Item.Description);
                TempCSVBuffer.InsertEntry(LineNo, 3, Format(Item.Inventory));
            until Item.Next() = 0;

        TempCSVBuffer.SaveDataToBlob(TempBlob, ';'); // 2nd parameter is a CSVFieldSeparator we can add ';' also
        TempBlob.CreateInStream(InStreamL);
        FileName := 'Item.csv';
        DownloadFromStream(InStreamL, '', '', '', FileName);
    end;

    local procedure ImportItemsFromCSVUsingCSVBuffer()
    var
        Item: Record Item;
        TempCSVBuffer: Record "CSV Buffer" temporary;
        FilePath: Text;
        InFileStream: InStream;
    begin
        if UploadIntoStream('Select File..', '', '', FilePath, InFileStream) then begin
            TempCSVBuffer.DeleteAll();
            TempCSVBuffer.LoadDataFromStream(InFileStream, ',');//CSVFieldSeparator
            if TempCSVBuffer.FindSet() then
                repeat
                    if TempCSVBuffer."Field No." = 1 then
                        Item.Init();

                    case TempCSVBuffer."Field No." of
                        1:
                            Item.Validate("No.", TempCSVBuffer.Value);
                        2:
                            Item.Validate(Description, TempCSVBuffer.Value);
                        3:
                            if not Item.Insert() then
                                Item.Modify();
                    end;
                until TempCSVBuffer.Next() = 0;
        end;
    end;

    local procedure ImportSalesOrderDocumentsFromTextUsingCSVBuffer()
    var
        Item: Record Item;
        TempCSVBuffer: Record "CSV Buffer";
        FilePath: Text;
        InFileStream: InStream;
    begin
        //In this case actually reading Text file with sales header and lines details field separated by ';' char with below format
        //H;SalesHeaderNumber;SellToCustomerNumber
        //L;LineType;LineNo;Qty;UnitPrice
        //L;LineType;LineNo;Qty;UnitPrice

        if UploadIntoStream('Select File..', '', '', FilePath, InFileStream) then begin
            TempCSVBuffer.DeleteAll();
            TempCSVBuffer.LoadDataFromStream(InFileStream, ';');//CSVFieldSeparator

            //Load only records whose field no is 1
            TempCSVBuffer.SetRange("Field No.", 1);
            if TempCSVBuffer.FindSet() then
                repeat
                    case TempCSVBuffer.Value of
                        'H':
                            CreateSalesOrder(TempCSVBuffer."Line No.");
                        'L':
                            CreateSalesLine(TempCSVBuffer."Line No.");
                    end;
                until TempCSVBuffer.Next() = 0;
        end;
    end;

    local procedure CreateSalesOrder(CSVBufferLineNo: Integer)
    var
        TempCSVBuffer: Record "CSV Buffer";
    begin
        Clear(SalesHeader);
        TempCSVBuffer.SetRange("Line No.", CSVBufferLineNo);
        if TempCSVBuffer.FindSet() then
            repeat
                case TempCSVBuffer."Field No." of
                    1:
                        begin
                            SalesHeader.Init();
                            SalesHeader.Validate("Document Type", SalesHeader."Document Type"::Order);
                        end;
                    2:
                        SalesHeader.Validate("No.", TempCSVBuffer.Value);
                    3:
                        begin
                            SalesHeader.Validate("Sell-to Customer No.", TempCSVBuffer.Value);
                            SalesHeader.Insert(true);
                        end;
                end;
            until TempCSVBuffer.Next() = 0;
    end;

    local procedure CreateSalesLine(CSVBufferLineNo: Integer)
    var
        SalesLine: Record "Sales Line";
        TempCSVBuffer: Record "CSV Buffer";
        SalesLineNo: Integer;
    begin
        TempCSVBuffer.SetRange("Line No.", CSVBufferLineNo);
        if TempCSVBuffer.FindSet() then
            repeat
                case TempCSVBuffer."Field No." of
                    1:
                        begin
                            SalesLineNo := GetSalesLineNo();
                            SalesLine.Init();
                            SalesLine.Validate("Document Type", SalesHeader."Document Type");
                            SalesLine.Validate("Document No.", SalesHeader."No.");
                            SalesLine.Validate("Line No.", SalesLineNo);
                        end;
                    2:
                        if TempCSVBuffer.Value > '' then
                            Evaluate(SalesLine.Type, TempCSVBuffer.Value);
                    3:
                        SalesLine.Validate("No.", TempCSVBuffer.Value);
                    4:
                        if TempCSVBuffer.Value > '' then begin
                            Evaluate(SalesLine.Quantity, TempCSVBuffer.Value);
                            SalesLine.Validate(Quantity);
                        end;
                    5:
                        begin
                            if TempCSVBuffer.Value > '' then begin
                                Evaluate(SalesLine."Unit Price", TempCSVBuffer.Value);
                                SalesLine.Validate("Unit Price");
                            end;
                            SalesLine.Insert(true);
                        end;
                end;
            until TempCSVBuffer.Next() = 0;
    end;

    local procedure GetSalesLineNo(): Integer
    var
        SalesLine: Record "Sales Line";
    begin
        SalesLine.SetRange("Document Type", SalesHeader."Document Type");
        SalesLine.SetRange("Document No.", SalesHeader."No.");
        if SalesLine.FindLast() then
            Exit(SalesLine."Line No." + 10000);

        exit(10000);
    end;

    var
        SalesHeader: Record "Sales Header";
}