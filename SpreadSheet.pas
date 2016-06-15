unit SpreadSheet;

interface

uses
  Graphics, Contnrs, XMLDoc, XMLIntf, Classes, Types;

type  
  THorizontalAlignment = (taLeft, taCenter, taRight);
  TVerticalAlignment = (taTop, taMiddle, taBottom);

  TAlignment = class
    Horizontal: THorizontalAlignment;
    Vertical: TVerticalAlignment;
  end;

  TFillType = (ftNone, ftSolid);
  
  TFill = class
    FillType: TFillType;
    Background: TColor;
    Foreground: TColor;
    constructor Create(FillType: TFillType; Background, Foreground: TColor);
  end;
  
  TBorderThickness = (btNone, btThin, btThick);
  
  TBorder = class
    Color: TColor;
    Thickness: TBorderThickness;
    constructor Create(Color: TColor; Thickness: TBorderThickness);
  end;
  
  TBorders = class
  private
    fLeft: TBorder;
    fRight: TBorder;
    fTop: TBorder;
    fBottom: TBorder;
  public
    constructor Create(Left, Right, Top, Bottom: TBorder);
    property Left: TBorder read fLeft write fLeft;
    property Right: TBorder read fRight write fRight;
    property Top: TBorder read fTop write fTop;
    property Bottom: TBorder read fBottom write fBottom;
  end;
  
  TFormattingRule = class
  private
    fID: string;
    fName: string;
    fFont: TFont;
    fFill: TFill;
    fBorders: TBorders;
    fFormatStr: string;
    fAlignment: TAlignment;
    fHeight: Double;
    fWidth: Double;
  public
    constructor Create(Name: string);
    destructor Destroy; override;
    procedure SetFont(Name: string; Size: Integer; Color: TColor; Style: TFontStyles);
    
    property Alignment: TAlignment read fAlignment write fAlignment;
    property ID: string read fID;
    property Name: string read fName;
    property Borders: TBorders read fBorders write fBorders;
    property Font: TFont read fFont;
    property Fill: TFill read fFill write fFill;
    property FormatStr: string read fFormatStr write fFormatStr;
    property Height: Double read fHeight write fHeight;
    property Width: Double read fWidth write fWidth;
  end;

  TFormattingRules = class
  private
    fList: TObjectList;
    function GetCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;

    procedure AddRule(Rule: TFormattingRule);
    property Count: Integer read GetCount;
  end;
  
  TCell = record
    Value: Variant;
    Formula: string;
    Format: TFormattingRule;
  end;

  TCellArray = array of TCell;

  TFontList = class(TObjectList)
  private
    function FontEquals(Font1, Font2: TFont): Boolean;
  public
    function IndexOf(Font: TFont): Integer; reintroduce;
  end;
  
  TStylesWriter = class
  private
    fNumFmts: TStringList;
    fFonts: TFontList;
    fFills: TObjectList;
    fBorders: TObjectList;
    fCellStyleXfs: TObjectList;
    fCellXfs: TObjectList;
    fStyles: TFormattingRules;
    procedure PrepareLists;
    procedure SaveFills(FillsNode: IXMLNode);
    procedure SaveBorders(BordersNode: IXMLNode);
  public
    constructor Create(Styles: TFormattingRules);
    destructor Destroy; override;
    
    procedure SaveToFile(FileName: string);
  end;
  
  TCellCoord = record
    Col: Integer;
    Row: Integer;
  end;
  
  TRange = record
    Coord1: TCellCoord;
    Coord2: TCellCoord;
  end;
  
  TWorkSheet = class
  private
    fName: string;
    fFormatRules: TFormattingRules;
    fColCount: Integer;
    fRowCount: Integer;
    fRows: array of TCellArray;
    fColFormats: TCellArray;
    fMergedCells: array of TRange;
    fColWidths: array of Double;
    fRowHeights: array of Double;
    fPrintRows: string;
    fPrintCols: string;
    procedure SetColCount(const Value: Integer);
    
    procedure SetRowCapacity(Count: Integer);
    procedure SaveColumnData(ParentNode: IXMLNode);
    procedure SaveCellFormat(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
    procedure SaveFormulaCell(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
    procedure SaveSheetData(ParentNode: IXMLNode);
    function AddCell(RowNode: IXMLNode; Col, Row: Integer): IXMLNode;
    procedure SaveNumberCell(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
    procedure SaveDateCell(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
    procedure SaveStringCell(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
    procedure SaveStyles(FileName: string);
    procedure SaveMergedCellsData(ParentNode: IXMLNode);
  public
    constructor Create(Name: string);
    destructor Destroy; override;

    procedure SetColFormat(Format: TFormattingRule; Index: Integer);
    procedure SetColWidth(Index: Integer; Width: Double);
    procedure SetRowHeight(Index: Integer; Height: Double);
    procedure SetCellFormat(Format: TFormattingRule; Col, Row: Integer);
    function IsParentMergedCell(Col, Row: Integer): Boolean;
    procedure UpdateMergedCellsFormat(Format: TFormattingRule; Col, Row: Integer);
    procedure FillFormula(Formula: string; Col, Row: integer);
    procedure FillString(Str: string; Col, Row: Integer);
    procedure FillVariant(Value: Variant; Col, Row: Integer);
    procedure MergeCells(Col1, Row1, Col2, Row2: Integer);
    procedure SaveToFile(const FileName: string);
    procedure SetPrintTitle(Rows, Cols: string);
    
    property ColCount: Integer read fColCount write SetColCount;
    property Name: string read fName write fName;
  end;

  TWorkBook = class
  private
    fBaseDir: string;
    fWorkSheetsDir: string;
    fWorkSheets: TObjectList;
    procedure SetBaseDir(const Value: string);
    procedure UpdateWorkSheetDir;
    procedure SavePrintTitles(Doc: IXMLDocument);
  public
    constructor Create;
    destructor Destroy; override;

    procedure AddWorksheet(Worksheet: TWorkSheet);
    procedure SaveRels(FileName: string);
    procedure SaveToFile(FileName: string);

    property BaseDir: string read fBaseDir write SetBaseDir;
  end;

  TSpreadSheet = class
  private
    fFilePath: string;
    fWorkBook: TWorkBook;
    fWorkingDir: string;
    procedure SaveContentTypes(FileName: string);
    procedure SaveRels(FileName: string);
    procedure ArchiveFile(Dir, FilePath: string);
    procedure SetWorkingDir(const Value: string);
    procedure CleanDirs(Dir: string);
    procedure SaveDocProps(ParentDir: string);
  public
    constructor Create(FilePath: string);
    destructor Destroy; override;
    procedure SaveToFile;

    property WorkingDir: string read fWorkingDir write SetWorkingDir;
    property WorkBook: TWorkBook read fWorkBook;
  end;

function GenerateColIdx(Index: Integer): string;  

implementation

uses
  Variants, TypInfo, SysUtils, ZipForge, ShellAPI, Windows, Math;
  
const
  NAMESPACE_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  NAMESPACE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  NAMESPACE_CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
  NAMESPACE_PROPERTIES = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';

  RELATIONSHIP_TYPE_WORKBOOK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  RELATIONSHIP_TYPE_WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
  RELATIONSHIP_TYPE_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
  RELATIONSHIP_TYPE_CORE_PROPS = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
  RELATIONSHIP_TYPE_EXT_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
  GROW_DELTA = 10;

function GenerateColIdx(Index: Integer): string;
const
  LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
begin
  Result := '';
  Index := Index - 1;
  while (Index > -1) do begin
    Result := Chr((Index mod 26) + 65) + Result;
    Index := (Index div 26) - 1;
  end;
end;
  
{ TReportXLSX }

constructor TWorkSheet.Create(Name: string);
begin
  fName := Name;
  fFormatRules := TFormattingRules.Create;
end;

destructor TWorkSheet.Destroy;
var
  i: Integer;
begin
  fFormatRules.Free;
  
  for i := 0 to Length(fRows) - 1 do begin
    SetLength(fRows[i], 0);
  end;

  SetLength(fRows, 0);
  SetLength(fColFormats, 0);
  SetLength(fColWidths, 0);
  SetLength(fRowHeights, 0);
end;

procedure TWorkSheet.FillVariant(Value: Variant; Col, Row: Integer);
begin
  SetRowCapacity(Row);

  fRows[Row - 1][Col - 1].Value := Value;
end;

procedure TWorkSheet.FillFormula(Formula: string; Col, Row: integer);
begin
  SetRowCapacity(Row);

  fRows[Row - 1][Col - 1].Formula := Formula;
end;

procedure TWorkSheet.FillString(Str: string; Col, Row: Integer);
begin
  SetRowCapacity(Row);

  fRows[Row - 1][Col - 1].Value := Str;
end;

procedure TWorkSheet.MergeCells(Col1, Row1, Col2, Row2: Integer);
begin
  SetLength(fMergedCells, Length(fMergedCells) + 1); //Slow, but i'll refactor later
  fMergedCells[Length(fMergedCells) - 1].Coord1.Col := Col1;
  fMergedCells[Length(fMergedCells) - 1].Coord1.Row := Row1;
  fMergedCells[Length(fMergedCells) - 1].Coord2.Col := Col2;
  fMergedCells[Length(fMergedCells) - 1].Coord2.Row := Row2;
end;

procedure TWorkSheet.SaveToFile(const FileName: string);
var
  Document: IXMLDocument;
  WorkSheet: IXMLNode;
  
  Cols: IXMLNode;
  MergedCells: IXMLNode;
  SheetDataNode: IXMLNode;
begin
  SaveStyles(ExtractFilePath(FileName) + '..\styles.xml');
  
  Document := NewXMLDocument;
  Document.Encoding := 'UTF-8';
  Document.StandAlone := 'yes';
  Document.Options := [doNodeAutoIndent];

  WorkSheet := Document.AddChild('worksheet', NAMESPACE_MAIN);
  WorkSheet.SetAttributeNS('xmlns:r',  '', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

  Cols := WorkSheet.AddChild('cols');
  SaveColumnData(Cols);
  if (not Cols.HasChildNodes) then begin
    Document.DocumentElement.ChildNodes.Delete('cols');
  end;

  SheetDataNode := WorkSheet.AddChild('sheetData');
  SaveSheetData(SheetDataNode);

  //Save merged cells. Needs to be after 'sheetData'
  MergedCells := WorkSheet.AddChild('mergeCells');
  SaveMergedCellsData(MergedCells);
  if (not MergedCells.HasChildNodes) then begin
    Document.DocumentElement.ChildNodes.Delete('mergeCells');
  end;

  Document.SaveToFile(FileName);
end;

procedure TWorkSheet.SaveStyles(FileName: string);
var
  Writer: TStylesWriter;
begin
  Writer := TStylesWriter.Create(fFormatRules);
  try
    Writer.SaveToFile(FileName);
  finally
    Writer.Free;
  end;
end;

procedure TWorkSheet.SaveColumnData(ParentNode: IXMLNode);
var
  i: Integer;
  Col: IXMLNode;
begin
  for i := 0 to fColCount - 1 do begin
    if (fColWidths[i] <> 0) then begin
      Col := ParentNode.AddChild('col');
      Col.Attributes['min'] := IntToStr(i + 1);
      Col.Attributes['max'] := IntToStr(i + 1);
      Col.Attributes['width'] := fColWidths[i];
    end;
  end;
end;

procedure TWorkSheet.SaveSheetData(ParentNode: IXMLNode);
var
  RowNode: IXMLNode;
  CellNode: IXMLNode;
  Cell: TCell;
  CellType: Integer;
  i, j: Integer;
begin
  for i := 0 to fRowCount - 1 do begin
    RowNode := ParentNode.AddChild('row');
    if (fRowHeights[i] <> 0) then begin
      RowNode.Attributes['ht'] := fRowHeights[i + 1];
      RowNode.Attributes['customHeight'] := '1';
    end;
    for j := 0 to Length(fRows[i]) - 1 do begin
      Cell := fRows[i][j];
      CellType := VarType(fRows[i][j].Value) and varTypeMask;
      
      if (fRows[i][j].Formula <> '') then begin
        CellNode := AddCell(RowNode, j, i);
        SaveFormulaCell(Cell, CellNode, j, i);
      end else if (CellType in [varSmallint, varInteger, varBoolean, varShortInt, varByte, varWord, varLongWord, varInt64]) then begin
        CellNode := AddCell(RowNode, j, i);
        SaveNumberCell(Cell, CellNode, j, i);
      end else if (CellType in [varSingle, varDouble, varCurrency]) then begin
        CellNode := AddCell(RowNode, j, i);
        SaveNumberCell(Cell, CellNode, j, i);
      end else if (CellType in [varDate]) then begin
        CellNode := AddCell(RowNode, j, i);
        SaveDateCell(Cell, CellNode, j, i);
      end else if (CellType = varString) then begin
        CellNode := AddCell(RowNode, j, i);
        SaveStringCell(Cell, CellNode, j, i);
      end else if (fRows[i][j].Format <> nil) then begin
        CellNode := AddCell(RowNode, j, i);
        SaveCellFormat(Cell, CellNode, j, i);
      end;
    end;
  end;
end;

function TWorkSheet.AddCell(RowNode: IXMLNode; Col, Row: Integer): IXMLNode;
begin
  if (RowNode <> nil) then begin
    Result := RowNode.AddChild('c');
    Result.Attributes['r'] := GenerateColIdx(Col + 1) + IntToStr(Row + 1);
  end;
end;

procedure TWorkSheet.SaveCellFormat(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
begin
  if (Cell.Format <> nil) then begin
    CellNode.Attributes['s'] := Cell.Format.ID;
  end else if (fColFormats[Col].Format <> nil) then begin
    CellNode.Attributes['s'] := fColFormats[Col].Format.ID;
  end;
end;

procedure TWorkSheet.SaveFormulaCell(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
var
  Node: IXMLNode;
begin
  if (CellNode <> nil) then begin
    SaveCellFormat(Cell, CellNode, Col, Row);
  
    Node := CellNode.AddChild('f');
    Node.NodeValue := VarToStr(fRows[Row][Col].Formula);
  end;
end;

procedure TWorkSheet.SaveNumberCell(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
var
  Node: IXMLNode;
begin
  if (CellNode <> nil) then begin
    SaveCellFormat(Cell, CellNode, Col, Row);

    CellNode.Attributes['t'] := 'n';
    Node := CellNode.AddChild('v');
    Node.NodeValue := fRows[Row][Col].Value;
  end;
end;

procedure TWorkSheet.SaveDateCell(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
var
  Date: Extended;
  Node: IXMLNode;
begin
  if (CellNode <> nil) then begin
    SaveCellFormat(Cell, CellNode, Col, Row);
  
    Node := CellNode.AddChild('v');
    Date := VarToDateTime(fRows[Row][Col].Value);
    Node.NodeValue := Date;
  end;
end;

procedure TWorkSheet.SaveStringCell(Cell: TCell; CellNode: IXMLNode; Col, Row: Integer);
var
  Node: IXMLNode;
begin
  if (CellNode <> nil) then begin
    CellNode.Attributes['t'] := 'inlineStr';
    SaveCellFormat(Cell, CellNode, Col, Row);
  
    Node := CellNode.AddChild('is').AddChild('t');
    Node.NodeValue := VarToStr(fRows[Row][Col].Value);
  end;
end;

procedure TWorkSheet.SaveMergedCellsData(ParentNode: IXMLNode);
var
  MergedCell: IXMLNode;
  i: Integer;
begin
  for i := 0 to Length(fMergedCells) - 1 do begin
    MergedCell := ParentNode.AddChild('mergeCell');
    MergedCell.Attributes['ref'] := 
      Format('%s%d:%s%d', 
        [
          GenerateColIdx(fMergedCells[i].Coord1.Col), fMergedCells[i].Coord1.Row,
          GenerateColIdx(fMergedCells[i].Coord2.Col), fMergedCells[i].Coord2.Row
        ]
      );
  end;
end;


procedure TWorkSheet.SetCellFormat(Format: TFormattingRule; Col, Row: Integer);
begin
  if (((Col - 1) < fColCount) and ((Row - 1) < fRowCount)) then begin
    if (IsParentMergedCell(Col, Row)) then begin
      UpdateMergedCellsFormat(Format, Col, Row);
    end;
    fRows[Row - 1][Col - 1].Format := Format;
    fFormatRules.AddRule(Format);
  end;
end;

function TWorkSheet.IsParentMergedCell(Col, Row: Integer): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to Length(fMergedCells) - 1 do begin
    if ((TRange(fMergedCells[i]).Coord1.Col = Col) and (TRange(fMergedCells[i]).Coord1.Row = Row)) then begin
      Result := True;
      Break;
    end;
  end;
end;

procedure TWorkSheet.UpdateMergedCellsFormat(Format: TFormattingRule; Col, Row: Integer);
var
  Range: TRange;
  i, j: Integer;
begin
  for i := 0 to Length(fMergedCells) - 1 do begin
    if ((TRange(fMergedCells[i]).Coord1.Col = Col) and (TRange(fMergedCells[i]).Coord1.Row = Row)) then begin
      Range := TRange(fMergedCells[i]);
      Break;
    end;
  end;

  for i := Range.Coord1.Col to Range.Coord2.Col do begin
    for j := Range.Coord1.Row to Range.Coord2.Row do begin
      if ((i <> Col) and (j <> Col)) then begin
        SetCellFormat(Format, i, j);
      end;
    end;
  end;
end;

procedure TWorkSheet.SetColCount(const Value: Integer);
var
  i: Integer;
begin
  fColCount := Value;

  SetLength(fColFormats, Value);
  SetLength(fColWidths, Value);
  for i := 0 to Length(fRows) - 1 do begin
    SetLength(fRows[i], fColCount);
  end;
end;

procedure TWorkSheet.SetColFormat(Format: TFormattingRule; Index: Integer);
begin
  if ((Index - 1) < fColCount) then begin
    fColFormats[Index - 1].Format := Format;
    fFormatRules.AddRule(Format);
  end;
end;

procedure TWorkSheet.SetColWidth(Index: Integer; Width: Double);
begin
  if ((Index - 1) < fColCount) then begin
    fColWidths[Index - 1] := Width;
  end;
end;

procedure TWorkSheet.SetPrintTitle(Rows, Cols: string);
begin
  fPrintRows := Rows;
  fPrintCols := Cols;
end;

procedure TWorkSheet.SetRowHeight(Index: Integer; Height: Double);
begin
  if ((Index - 1) < fRowCount) then begin
    fRowHeights[Index - 1] := Height;
  end;
end;

procedure TWorkSheet.SetRowCapacity(Count: Integer);
var
  NewCount: Integer;
  i: Integer;
begin
  if (Count > fRowCount) then begin
    NewCount := ((Count div GROW_DELTA) + 1) * GROW_DELTA;
    SetLength(fRows, NewCount);
    SetLength(Self.fRowHeights, NewCount);
    for i := fRowCount to NewCount - 1 do begin
      SetLength(fRows[i], fColCount);
    end;
    fRowCount := NewCount;
  end;
end;

{ TFormattingRule }

constructor TFormattingRule.Create(Name: string);
begin
  fName := Name;
  fAlignment := TAlignment.Create;
  fFont := TFont.Create;
  fFont.Name := 'Arial';
  fFont.Size := 10;
end;

destructor TFormattingRule.Destroy;
begin
  fAlignment.Free;
  fFont.Free;
  fFill.Free;
end;

procedure TFormattingRule.SetFont(Name: string; Size: Integer; Color: TColor; Style: TFontStyles);
begin
  fFont.Name := Name;
  fFont.Size := Size;
  fFont.Color := Color;
  fFont.Style := Style;
end;

{ TFormattingRules }

procedure TFormattingRules.AddRule(Rule: TFormattingRule);
begin
  if (fList.IndexOf(Rule) < 0) then begin
    Rule.fID := IntToStr(Count);
    fList.Add(Rule);
  end;
end;

constructor TFormattingRules.Create;
begin
  fList := TObjectList.Create;
  
  AddRule(TFormattingRule.Create(''));
end;

destructor TFormattingRules.Destroy;
begin
  fList.Free;
end;

function TFormattingRules.GetCount: Integer;
begin
  Result := fList.Count;
end;

{ TSylesWriter }

constructor TStylesWriter.Create(Styles: TFormattingRules);
begin
  fStyles := Styles;

  fNumFmts := TStringList.Create;
  fFonts := TFontList.Create(False);
  fFills := TObjectList.Create(False); 
  fBorders := TObjectList.Create;
  fCellStyleXfs := TObjectList.Create;
  fCellXfs := TObjectList.Create;
end;

destructor TStylesWriter.Destroy;
begin
  fNumFmts.Free;
  fFonts.Free;
  fFills.Free;
  fBorders.Free;
  fCellStyleXfs.Free;
  fCellXfs.Free;
end;

procedure TStylesWriter.PrepareLists;
var
  i: Integer;
  Rule: TFormattingRule;
begin
  fBorders.Add(TBorders.Create(nil, nil, nil, nil));
  
  for i := 0 to fStyles.Count - 1 do begin
    Rule := TFormattingRule(fStyles.fList[i]);

    //Check fonts
    if ((TFormattingRule(fStyles.fList[i]).Font <> nil) and (fFonts.IndexOf(TFormattingRule(fStyles.fList[i]).Font) < 0)) then begin
      fFonts.Add(Rule.Font);
    end;

    //numStr
    if ((Rule.FormatStr <> '') and (fNumFmts.IndexOf(Rule.FormatStr) < 0)) then begin
      fNumFmts.Add(Rule.FormatStr);
    end;

    //fill
    if (Assigned(Rule.Fill)) then begin
      if (fFills.IndexOf(Rule.Fill) < 0) then begin
        fFills.Add(Rule.Fill);
      end;
    end;

    //borders
    if (Assigned(Rule.fBorders)) then begin
      if (fBorders.IndexOf(Rule.Borders) < 0) then begin
        fBorders.Add(Rule.Borders);
      end;
    end;
  end;
end;



procedure TStylesWriter.SaveToFile(FileName: string);
var
  Doc: IXMLDocument;
  NumFmts: IXMLNode;
  Fonts: IXMLNode;
  Fills: IXMLNode;
  Borders: IXMLNode;
//  CellStyleXFs: IXMLNode;
  CellXFs: IXMLNode;
  Node: IXMLNode;
  Alignment: IXMLNode;
  Rule: TFormattingRule;
  i: Integer;
begin
  if (fStyles <> nil) then begin
    PrepareLists;
    
    Doc := NewXMLDocument;
    Doc.Encoding := 'UTF-8';
    Doc.StandAlone := 'yes';
//    Doc.Options := [doNodeAutoIndent];

    Doc.AddChild('styleSheet', NAMESPACE_MAIN);

    //Saving number formatting rules
    NumFmts := Doc.DocumentElement.AddChild('numFmts');
    NumFmts.Attributes['count'] := fNumFmts.Count;

    for i := 0 to fNumFmts.Count - 1 do begin
      Node := NumFmts.AddChild('numFmt');
      Node.Attributes['numFmtId'] := IntToStr(i);
      Node.Attributes['formatCode'] := fNumFmts[i];
    end;

    //Saving fonts
    Fonts := Doc.DocumentElement.AddChild('fonts');
    Fonts.Attributes['count'] := fFonts.Count;

    for i := 0 to fFonts.Count - 1 do begin
      Node := Fonts.AddChild('font');
      if (fsBold in TFont(fFonts[i]).Style) then begin
        Node.AddChild('b');
      end;

      Node.AddChild('sz').Attributes['val'] := TFont(fFonts[i]).Size;
      Node.AddChild('name').Attributes['val'] := TFont(fFonts[i]).Name;
      
      if (TFont(fFonts[i]).Color <> clBlack) then begin
        Node.AddChild('color').Attributes['rgb'] := IntToHex(TFont(fFonts[i]).Color, 8);
      end;
    end;

    //Save fills
    Fills := Doc.DocumentElement.AddChild('fills');
    Fills.Attributes['count'] := fFills.Count;
    SaveFills(Fills);

    //Save borders
    Borders := Doc.DocumentElement.AddChild('borders');
    Borders.Attributes['count'] := fBorders.Count;
    SaveBorders(Borders);

    //Saving cell XFs
    CellXFs := Doc.DocumentElement.AddChild('cellXfs');
    CellXFs.Attributes['count'] := fStyles.fList.Count;

    for i := 0 to fStyles.fList.Count - 1 do begin
      Rule := TFormattingRule(fStyles.fList[i]);
      Node := CellXFs.AddChild('xf');
      if (Rule.FormatStr <> '') then begin
        Node.Attributes['numFmtId'] := fNumFmts.IndexOf(Rule.FormatStr);
      end;

      if ((Rule.Font <> nil) and (Rule.Font.Name <> '')) then begin
        Node.Attributes['fontId'] := fFonts.IndexOf(Rule.Font);
      end;

      if ((Rule.Alignment.Horizontal <> taLeft) or (Rule.Alignment.Vertical <> taTop)) then begin
        Alignment := Node.AddChild('alignment');
        if (Rule.Alignment.Horizontal <> taLeft) then begin
          case (Rule.Alignment.Horizontal) of
            taCenter: Alignment.Attributes['horizontal'] := 'center';
            taRight: Alignment.Attributes['horizontal'] := 'right';
          end;
        end;
        if (Rule.Alignment.Vertical <> taBottom) then begin
          case (Rule.Alignment.Vertical) of
            taTop: Alignment.Attributes['vertical'] := 'top';
            taMiddle: Alignment.Attributes['vertical'] := 'center';
            taBottom: Alignment.Attributes['vertical'] := 'bottom';
          end;
        end;
      end;

      if (Assigned(Rule.Fill)) then begin
        Node.Attributes['fillId'] := fFills.IndexOf(Rule.Fill);
      end;

      if (Assigned(Rule.Borders)) then begin
        Node.Attributes['borderId'] := fBorders.IndexOf(Rule.Borders);
      end;
    end;
    
    Doc.SaveToFile(FileName);

    Doc := nil;
  end;
end;

procedure TStylesWriter.SaveFills(FillsNode: IXMLNode);
var
  i: Integer;
  Fill: IXMLNode;
  Node: IXMLNode;
  DummyFill: TFill;
begin
  DummyFill := TFill.Create(ftNone, $1FFFFFFF, $1FFFFFFF);
  try
    fFills.Insert(0, DummyFill); //Excel seems to need these empty fills;
    fFills.Insert(0, DummyFill);
    
    for i := 0 to fFills.Count - 1 do begin
      Fill := FillsNode.AddChild('fill');
      if (TFill(fFills[i]).FillType <> ftNone) then begin
        Node := Fill.AddChild('patternFill');
        Node.Attributes['patternType'] := 'solid';
        Node.AddChild('fgColor').Attributes['rgb'] := IntToHex(TFill(fFills[i]).Foreground, 8);
        Node.AddChild('bgColor').Attributes['rgb'] := IntToHex(TFill(fFills[i]).Background, 8);
      end;
    end;
  finally
    DummyFill.Free;
  end;
end;

procedure TStylesWriter.SaveBorders(BordersNode: IXMLNode);
var
  i: Integer;
  Borders: TBorders;
  BorderNode: IXMLNode;
  Node: IXMLNode;
begin
  {
  <border>
      <left style="thin">
        <color indexed="64"/>
      </left>
      <right style="thin">
        <color indexed="64"/>
      </right>
      <top style="thin">
        <color indexed="64"/>
      </top>
      <bottom style="thin">
        <color indexed="64"/>
      </bottom>
      <diagonal/>
    </border>}
  for i := 0 to fBorders.Count - 1 do begin
    BorderNode := BordersNode.AddChild('border');
    Borders := TBorders(fBorders[i]);
    if (Assigned(Borders.Left)) then begin
      Node := BorderNode.AddChild('left');
      Node.Attributes['style'] := 'thin';
      Node.AddChild('color').Attributes['rgb'] := IntToHex(Borders.Left.Color, 8);
    end;
    if (Assigned(Borders.Right)) then begin
      Node := BorderNode.AddChild('right');
      Node.Attributes['style'] := 'thin';
      Node.AddChild('color').Attributes['rgb'] := IntToHex(Borders.Right.Color, 8);
    end;
    if (Assigned(Borders.Top)) then begin
      Node := BorderNode.AddChild('top');
      Node.Attributes['style'] := 'thin';
      Node.AddChild('color').Attributes['rgb'] := IntToHex(Borders.Top.Color, 8);
    end;
    if (Assigned(Borders.Bottom)) then begin
      Node := BorderNode.AddChild('bottom');
      Node.Attributes['style'] := 'thin';
      Node.AddChild('color').Attributes['rgb'] := IntToHex(Borders.Bottom.Color, 8);
    end;
  end;
end;

{ TFill }

constructor TFill.Create(FillType: TFillType; Background, Foreground: TColor);
begin
  Self.FillType := FillType;
  Self.Background := Background;
  Self.Foreground := Foreground;
end;

{ TBorder }

constructor TBorder.Create(Color: TColor; Thickness: TBorderThickness);
begin
  Self.Color := Color;
  Self.Thickness := Thickness;
end;

{ TBorders }

constructor TBorders.Create(Left, Right, Top, Bottom: TBorder);
begin
  fLeft := Left;
  fRight := Right;
  fTop := Top;
  fBottom := Bottom;
end;

{ TSpreadSheet }

procedure TSpreadSheet.ArchiveFile(Dir, FilePath: string);
var
  Archiver: TZipForge;
begin
  Archiver := TZipForge.Create(nil);
  try
    Archiver.BaseDir := Dir;
    Archiver.FileName := FilePath;
    Archiver.FileMasks.Text := '*.*';
    Archiver.CompressionLevel := clMax;
    Archiver.OpenArchive;
    Archiver.AddFiles;
    Archiver.CloseArchive;
  finally
    Archiver.Free;
  end;
end;

procedure TSpreadSheet.CleanDirs(Dir: string);
var
  FileOp: TSHFileOpStruct;
begin
  FillChar(FileOp, SizeOf(FileOp), 0);
  FileOp.wFunc := FO_DELETE;
  FileOp.pFrom := PChar(Dir + #0);//double zero-terminated
  FileOp.fFlags := FOF_SILENT or FOF_NOERRORUI or FOF_NOCONFIRMATION;
  SHFileOperation(FileOp);
end;

constructor TSpreadSheet.Create(FilePath: string);
begin
  fFilePath := FilePath;
  fWorkBook := TWorkBook.Create;
end;

destructor TSpreadSheet.Destroy;
begin
  fWorkBook.Free;
end;

procedure TSpreadSheet.SaveContentTypes(FileName: string);
const
  CONTENT_TYPE_RELATIONSHIPS = 'application/vnd.openxmlformats-package.relationships+xml';
  CONTENT_TYPE_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
  CONTENT_TYPE_WORKBOOK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  CONTENT_TYPE_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
  CONTENT_TYPE_EXTENDED_PROPS = 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
  CONTENT_TYPE_CORE_PROPS = 'application/vnd.openxmlformats-package.core-properties+xml';
var
  Doc: IXMLDocument;
  Node: IXMLNode;
  i: Integer;
begin
  Doc := NewXMLDocument;
  Doc.Encoding := 'UTF-8';
  Doc.StandAlone := 'yes';
//  Doc.Options := [doNodeAutoIndent];

  Doc.AddChild('Types', NAMESPACE_CONTENT_TYPES);

  //Rels node;
  Node := Doc.DocumentElement.AddChild('Default');
  Node.Attributes['Extension'] := 'rels';
  Node.Attributes['ContentType'] := CONTENT_TYPE_RELATIONSHIPS;

  Node := Doc.DocumentElement.AddChild('Default');
  Node.Attributes['Extension'] := 'xml';
  Node.Attributes['ContentType'] := 'application/xml';
  
  //Workbook node;
  Node := Doc.DocumentElement.AddChild('Override');
  Node.Attributes['PartName'] := '/workbook.xml';
  Node.Attributes['ContentType'] := CONTENT_TYPE_WORKBOOK;

  //Styles node;
  Node := Doc.DocumentElement.AddChild('Override');
  Node.Attributes['PartName'] := '/styles.xml';
  Node.Attributes['ContentType'] := CONTENT_TYPE_STYLES;

  //Author node;
  Node := Doc.DocumentElement.AddChild('Override');
  Node.Attributes['PartName'] := '/docProps/core.xml';
  Node.Attributes['ContentType'] := CONTENT_TYPE_CORE_PROPS;

  //Generator app node;
  Node := Doc.DocumentElement.AddChild('Override');
  Node.Attributes['PartName'] := '/docProps/app.xml';
  Node.Attributes['ContentType'] := CONTENT_TYPE_EXTENDED_PROPS;

  //Worksheets
  for i := 0 to fWorkBook.fWorkSheets.Count - 1 do begin
    Node := Doc.DocumentElement.AddChild('Override');
    Node.Attributes['PartName'] := '/sheets/sheet' + IntToStr(i + 1) + '.xml';
    Node.Attributes['ContentType'] := CONTENT_TYPE_WORKSHEET;
  end;
  
  Doc.SaveToFile(FileName);
  Doc := nil;                         
end;

procedure TSpreadSheet.SaveRels(FileName: string);
var
  Doc: IXMLDocument;
  Node: IXMLNode;
begin
  Doc := NewXMLDocument;
  Doc.Encoding := 'UTF-8';
  Doc.StandAlone := 'yes';
//  Doc.Options := [doNodeAutoIndent];

  Doc.AddChild('Relationships', NAMESPACE_RELATIONSHIPS);

  Node := Doc.DocumentElement.AddChild('Relationship');
  Node.Attributes['Id'] := 'rId1';
  Node.Attributes['Type'] := RELATIONSHIP_TYPE_WORKBOOK;
  Node.Attributes['Target'] := 'workbook.xml';

  Node := Doc.DocumentElement.AddChild('Relationship');
  Node.Attributes['Id'] := 'rId2';
  Node.Attributes['Type'] := RELATIONSHIP_TYPE_CORE_PROPS;
  Node.Attributes['Target'] := 'docProps/core.xml';

  Node := Doc.DocumentElement.AddChild('Relationship');
  Node.Attributes['Id'] := 'rId3';
  Node.Attributes['Type'] := RELATIONSHIP_TYPE_EXT_PROPS;
  Node.Attributes['Target'] := 'docProps/app.xml';
  
  Doc.SaveToFile(FileName);
  Doc := nil;
end;
                    
procedure TSpreadSheet.SaveToFile;
var
  FileDir: string;
  RelsDir: string;
begin
  try
    if (fWorkingDir = '') then begin
      fWorkingDir := IncludeTrailingPathDelimiter(GetCurrentDir);
    end;
    
    if (FileExists(fFilePath)) then begin
      Windows.DeleteFile(PChar(fFilePath));
    end;

    FileDir := ExtractFileName(fFilePath);
    FileDir := ChangeFileExt(FileDir, '');
    ForceDirectories(fWorkingDir + FileDir);
    
    RelsDir := fWorkingDir + FileDir + '\_rels\'; 
    ForceDirectories(RelsDir);

    SaveRels(RelsDir + '.rels');
    SaveContentTypes(fWorkingDir + FileDir + '\[Content_Types].xml');

    SaveDocProps(fWorkingDir + FileDir);
    
    fWorkBook.BaseDir := fWorkingDir + FileDir;
    fWorkBook.SaveToFile(fWorkingDir + FileDir + '\workbook.xml');
    fWorkBook.SaveRels(RelsDir + 'workbook.xml.rels');

    ArchiveFile(fWorkingDir + FileDir, fFilePath);
  finally
    CleanDirs(fWorkingDir + FileDir);
  end;
end;

procedure TSpreadSheet.SaveDocProps(ParentDir: string);
var
  PropsDir: string;
  Doc: IXMLDocument;
  Node: IXMLNode;
begin
  PropsDir := IncludeTrailingPathDelimiter(ParentDir) + 'docProps';
  ForceDirectories(PropsDir);

  Doc := NewXMLDocument;
  Doc.Encoding := 'UTF-8';
  Doc.StandAlone := 'yes';
//  Doc.Options := [doNodeAutoIndent];

  Doc.AddChild('cp:coreProperties');
  Doc.DocumentElement.Attributes['xmlns:cp'] := 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
  Doc.DocumentElement.Attributes['xmlns:dc'] := 'http://purl.org/dc/elements/1.1/';
  Doc.DocumentElement.Attributes['xmlns:dcterms'] := 'http://purl.org/dc/terms/';
  Doc.DocumentElement.Attributes['xmlns:dcmitype'] := 'http://purl.org/dc/dcmitype/';
  Doc.DocumentElement.Attributes['xmlns:xsi'] := 'http://www.w3.org/2001/XMLSchema-instance';

  Node := Doc.DocumentElement.AddChild('dc:creator');
  Node.NodeValue := 'dsContabilitate 2016';
  
  Node := Doc.DocumentElement.AddChild('dc:description');
  Node.NodeValue := 'Copyright 2016 SC D-Soft SRL; Parts copyright 2016 Danciu Alexandru Vasile';
  
  Doc.SaveToFile(PropsDir + '\core.xml');
  Doc := nil;

  Doc := NewXMLDocument;
  Doc.Encoding := 'UTF-8';
  Doc.StandAlone := 'yes';
//  Doc.Options := [doNodeAutoIndent];

  Doc.AddChild('Properties', NAMESPACE_PROPERTIES);
  Doc.DocumentElement.Attributes['xmlns:vt'] := 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

  Node := Doc.DocumentElement.AddChild('Application');
  Node.NodeValue := 'dsContabilitate 2016';

  Doc.SaveToFile(PropsDir + '\app.xml');
end;

procedure TSpreadSheet.SetWorkingDir(const Value: string);
begin
  fWorkingDir := IncludeTrailingPathDelimiter(Value);
end;

{ TWorkBook }

procedure TWorkBook.AddWorksheet(Worksheet: TWorkSheet);
begin
  if (fWorkSheets.IndexOf(Worksheet) < 0) then begin
    fWorkSheets.Add(Worksheet);
  end;
end;

constructor TWorkBook.Create;
begin
  fWorkSheets := TObjectList.Create;
end;

destructor TWorkBook.Destroy;
begin
  fWorkSheets.Free;
end;

procedure TWorkBook.SaveRels(FileName: string);
var
  i: Integer;
  Doc: IXMLDocument;
  Node: IXMLNode;
begin
  Doc := NewXMLDocument;
  Doc.Encoding := 'UTF-8';
  Doc.StandAlone := 'yes';
//  Doc.Options := [doNodeAutoIndent];

  Doc.AddChild('Relationships', NAMESPACE_RELATIONSHIPS);

  for i := 0 to fWorkSheets.Count - 1 do begin
    Node := Doc.DocumentElement.AddChild('Relationship');
    Node.Attributes['Id'] := 'rId' + IntToStr(i + 1);
    Node.Attributes['Type'] := RELATIONSHIP_TYPE_WORKSHEET;
    Node.Attributes['Target'] := 'sheets/sheet' + IntToStr(i + 1) + '.xml';
  end;

  Node := Doc.DocumentElement.AddChild('Relationship');
  Node.Attributes['Id'] := 'rId' + IntToStr(fWorkSheets.Count + 1);
  Node.Attributes['Type'] := RELATIONSHIP_TYPE_STYLES;
  Node.Attributes['Target'] := 'styles.xml';
  
  Doc.SaveToFile(FileName);
  Doc := nil;
end;

procedure TWorkBook.SaveToFile(FileName: string);
var
  Doc: IXMLDocument;
  Sheets: IXMLNode;
  Node: IXMLNode;
  i: Integer;
begin
  for i := 0 to fWorkSheets.Count - 1 do begin
    TWorkSheet(fWorkSheets[i]).SaveToFile(IncludeTrailingPathDelimiter(fWorkSheetsDir) + 'sheet' + IntToStr(i + 1) + '.xml');
  end;

  //Save workbook file;
  Doc := NewXMLDocument;
  Doc.Encoding := 'UTF-8';
  Doc.StandAlone := 'yes';
//  Doc.Options := [doNodeAutoIndent];

  Doc.AddChild('workbook', NAMESPACE_MAIN);
  Doc.DocumentElement.SetAttributeNS('xmlns:r',  '', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

  Sheets := Doc.DocumentElement.AddChild('sheets');
  
  for i := 0 to fWorkSheets.Count - 1 do begin
    Node := Sheets.AddChild('sheet');
    Node.Attributes['name'] := TWorkSheet(fWorkSheets[i]).Name;
    Node.Attributes['sheetId'] := IntToStr(i + 1);
    Node.Attributes['r:id'] := 'rId' + IntToStr(i + 1);
  end;

  SavePrintTitles(Doc);

  Doc.SaveToFile(FileName);
  Doc := nil;
end;

procedure TWorkBook.SetBaseDir(const Value: string);
begin
  fBaseDir := Value;
  UpdateWorkSheetDir;
end;

procedure TWorkBook.UpdateWorkSheetDir;
begin
  fWorkSheetsDir := IncludeTrailingPathDelimiter(fBaseDir) + 'sheets\';
  try
    CreateDir(fWorkSheetsDir);
  except 
  end;
end;

procedure TWorkBook.SavePrintTitles(Doc: IXMLDocument);
var
  DefinedNames: IXMLNode;
  Node: IXMLNode;
  PrintTitles: string;
  i: Integer;
begin
  for i := 0 to fWorkSheets.Count - 1 do begin
    if ((TWorkSheet(fWorkSheets[i]).fPrintCols <> '') and (TWorkSheet(fWorkSheets[i]).fPrintRows <> '')) then begin
      PrintTitles := 
        Format('''%s''!%s,''%s''!%s', 
        [
          TWorkSheet(fWorkSheets[i]).Name, TWorkSheet(fWorkSheets[i]).fPrintCols, 
          TWorkSheet(fWorkSheets[i]).Name, TWorkSheet(fWorkSheets[i]).fPrintRows
        ]
      );
    end else begin
      if (TWorkSheet(fWorkSheets[i]).fPrintCols <> '') then begin
        PrintTitles := Format('''%s''!%s', [TWorkSheet(fWorkSheets[i]).Name, TWorkSheet(fWorkSheets[i]).fPrintCols]);
      end;
      
      if (TWorkSheet(fWorkSheets[i]).fPrintRows <> '') then begin
        PrintTitles := Format('''%s''!%s', [TWorkSheet(fWorkSheets[i]).Name, TWorkSheet(fWorkSheets[i]).fPrintRows]);
      end;
    end;
    
    if (PrintTitles <> '') then begin
      DefinedNames := Doc.DocumentElement.AddChild('definedNames');
      Node := DefinedNames.AddChild('definedName');
      Node.Attributes['name'] := '_xlnm.Print_Titles';
      Node.Attributes['localSheetId'] := i;
      Node.Text := PrintTitles;
    end;
  end;
end;

{ TFontList }

function TFontList.FontEquals(Font1, Font2: TFont): Boolean;
begin
  Result := 
    (Font1.Name = Font2.Name) and (Font1.Size = Font2.Size) and 
    (Font1.Style = Font2.Style) and (Font1.Color = Font2.Color);
end;

function TFontList.IndexOf(Font: TFont): Integer;
var
  i: Integer;
  CurrFont: TFont;
begin
  Result := -1;
  for i := 0 to Count - 1 do begin
    CurrFont := TFont(Get(i));
    if (FontEquals(CurrFont, Font)) then begin
      Result := i;
      Break;
    end;
  end;
end;

end.
