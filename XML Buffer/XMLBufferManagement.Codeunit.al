codeunit 50100 "XML Buffer Management"
{
    trigger OnRun()
    begin
        //ExportItemToXMLUsingXMLBuffer(); // Simple
        //ExportSalesOrderToXMLUsingXMLBuffer(); // Complex

        ImportItemsFromXMLUsingXMLBuffer();
    end;

    local procedure ExportItemToXMLUsingXMLBuffer()
    var
        Item: Record Item;
        TempXMLBuffer: Record "XML Buffer" temporary;
    begin
        TempXMLBuffer.AddGroupElement('Items');
        Item.SetLoadFields("No.", Description, Inventory);// Load only required fields
        Item.SetAutoCalcFields(Inventory);//Set Calcfields outside the loop
        if Item.FindSet() then
            repeat
                TempXMLBuffer.AddGroupElement('Item');
                TempXMLBuffer.AddAttribute('No', Item."No.");
                TempXMLBuffer.AddElement('Description', Item.Description);
                TempXMLBuffer.AddElement('Inventory', Format(Item.Inventory));
                TempXMLBuffer.GetParent();
            until Item.Next() = 0;

        //TempXMLBuffer.Save('c:/Item.xml'); // OnPrem
        SaveXMLBufferToTempBlobAndDownload(TempXMLBuffer, 'Item.xml');//SAAS
    end;

    local procedure ExportSalesOrderToXMLUsingXMLBuffer()
    var
        SalesHdr: Record "Sales Header";
        SalesLine: Record "Sales Line";
        TempXMLBuffer: Record "XML Buffer" temporary;
        SalesOrdersEntryNo: Integer;
    begin
        //Save Sales Order Entry number from Temp XML Buffer
        SalesOrdersEntryNo := TempXMLBuffer.AddGroupElement('SalesOrders');
        SalesHdr.SetRange("Document Type", SalesHdr."Document Type"::Order);
        SalesHdr.SetLoadFields("No.", "Sell-to Customer No.");// Load only required fields
        if SalesHdr.FindSet() then
            repeat
                TempXMLBuffer.Get(SalesOrdersEntryNo); // Get SalesOrdersEntryNo Or Use TempXMLBuffer.GetParent
                TempXMLBuffer.AddGroupElement('Header');
                TempXMLBuffer.AddAttribute('OrderNumber', SalesHdr."No.");
                TempXMLBuffer.AddElement('SelltoCustomerNo', SalesHdr."Sell-to Customer No.");

                SalesLine.SetRange("Document Type", SalesHdr."Document Type");
                SalesLine.SetRange("Document No.", SalesHdr."No.");
                SalesLine.SetLoadFields(Type, "No.", Description, "Unit Price");// Load only required fields
                if SalesLine.FindSet() then
                    repeat
                        TempXMLBuffer.AddGroupElement('Line');
                        TempXMLBuffer.AddElement('Type', Format(SalesLine.Type));
                        TempXMLBuffer.AddElement('No', SalesLine."No.");
                        TempXMLBuffer.AddElement('Description', SalesLine.Description);
                        TempXMLBuffer.AddElement('UnitPrice', Format(SalesLine."Unit Price"));
                        TempXMLBuffer.GetParent();
                    until SalesLine.Next() = 0;
            until SalesHdr.Next() = 0;

        TempXMLBuffer.Get(SalesOrdersEntryNo); // Get SalesOrdersEntryNo or we have to use multiple TempXMLBuffer.GetParent
        SaveXMLBufferToTempBlobAndDownload(TempXMLBuffer, 'SalesOrders.xml');//SAAS
    end;

    local procedure SaveXMLBufferToTempBlobAndDownload(var TempXMLBuffer: Record "XML Buffer" temporary; FileName: Text)
    var
        XMLReader: Codeunit "XML Buffer Reader";
        TempBlob: Codeunit "Temp Blob";
        XMLDoc: XmlDocument;
        InStreamL: InStream;
    begin
        XMLReader.SaveToTempBlob(TempBlob, TempXMLBuffer);
        TempBlob.CreateInStream(InStreamL);
        XmlDocument.ReadFrom(InStreamL, XMLDoc);
        DownloadFromStream(InStreamL, '', '', '', FileName);
    end;

    local procedure ImportItemsFromXMLUsingXMLBuffer()
    var
        Item: Record Item;
        TempXMLBuffer: Record "XML Buffer" temporary;
        FilePath: Text;
        InStreamL: InStream;
    begin
        //TempXMLBuffer.Load('c:/Item.xml'); // OnPrem
        if UploadIntoStream('Select File..', '', '', FilePath, InStreamL) then
            TempXMLBuffer.LoadFromStream(InStreamL);

        if TempXMLBuffer.FindSet() then
            repeat
                if (TempXMLBuffer.Type = TempXMLBuffer.Type::Element) and (TempXMLBuffer.Name = 'Item') then
                    Item.Init();
                case TempXMLBuffer.Name of
                    'No':
                        Item.Validate("No.", TempXMLBuffer.Value);
                    'Description':
                        begin
                            Item.Validate(Description, TempXMLBuffer.Value);
                            if not Item.Insert() then
                                Item.Modify();
                        end;
                end;
            until TempXMLBuffer.Next() = 0;
    end;
}