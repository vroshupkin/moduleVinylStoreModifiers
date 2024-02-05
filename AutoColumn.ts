namespace AutoColumn
{   
    type Range = GoogleAppsScript.Spreadsheet.Range;
    

// export function CellInRange(x: number, y: number, range: Range)
// {
//   return !(
//     x < range.getColumn() ||
//     x > range.getColumn() + range.getWidth() - 1 || 
//     y < range.getRow() ||
//     y > range.getRow() + range.getHeight() - 1 
//   );
// }

 interface AutoColumnFuncParam {
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    /** Координаты заголовка с записью 'auto' */
    header_coords: {y: number, x: number},
    /** Координаты начало значений с которых начинается автозаполнение */
    values_coords: {y: number, x: number},
    /** size: columns - ширина таблица, в который ищется 'auto'. rows - глубина, на которой надо протянуть */
    size: {rows: number, columns: number}
  }

export function AutoColumnFunc(params: AutoColumnFuncParam, get_cell_type?: typeof AutoColumn.get_cell_type)
{
  const { sheet, values_coords, header_coords, size } = params;
  const header_range = sheet.getRange(header_coords.y, header_coords.x, 1, size.columns);
  const notes = header_range.getNotes()[0];
  
  if(!Array.isArray(notes)) { return; }

  const rows = size.rows - values_coords.y + 1;
  for (const [ i, note ] of notes.entries()) 
  { 
    if(i === 0)
    {
      continue;   
    }

    const x = values_coords.x + i;
    const note_type = (note + '').trim().toLowerCase();
    
    
    if(note_type.length > 4)
    {
      Logger.log(note_type);
    }

    if(note_type.includes('discogsFormat'.toLowerCase()))
    {
      const fillFormat = () => 
      {
        const param = note_type.slice(
          note_type.indexOf('(')+ 1,
          note_type.indexOf(')')
        );
      
        const source_x = Tools.calcXPosition(param);

        const insert_column = sheet.getRange(values_coords.y, x, size.rows);
        const source_column = sheet.getRange(values_coords.y, source_x, size.rows);

        Tools.fillTableByTable(insert_column, source_column, Discogs.formatHandle);
      };

      Tools.timestampDecorator(fillFormat, 'fillFormat')();
      
    }

    if(note_type !== 'auto')
    {
      continue;
    }

    const cell_type = get_cell_type?
      get_cell_type(values_coords.y, x, sheet): Constants.EnumCellType.DEFAULT;
    
    if(cell_type === Constants.EnumCellType.DEFAULT)
    {
      const pattern_cell = sheet.getRange(values_coords.y, x);
      const destination = sheet.getRange(values_coords.y, x, rows);
      pattern_cell.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); 
      
    }
    else if(cell_type === Constants.EnumCellType.CHECKBOX)
    {
      // Logger.log('Constants.EnumCellType.CHECKBOX');
      // !!! WARNING size_rows должен называется insert_y
      const insert_cell = sheet.getRange(size.rows, x);
      const check_box_data_validation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      
      insert_cell.setDataValidation(check_box_data_validation);

      insert_cell.setValue(false);
      
    }
    else if(cell_type === Constants.EnumCellType.DROPDOWN)
    {
      
      // Logger.log('Constants.EnumCellType.DROPDOWN');
      const pattern_cell = sheet.getRange(values_coords.y, x);
      // Logger.log(pattern_cell.getValue());
      // const data_validation = pattern_cell.getDataValidation();
      // Logger.log(data_validation + '');

      const destination = sheet.getRange(size.rows, x);
      pattern_cell.copyTo(destination);

      // destination.setDataValidation(data_validation);
      // destination.setValue(pattern_cell.getValue());

    }
  }
}


  /**
   * @param values_range - Ячейки откуда протягивается автозначения
   * Сохраняет в кеш по ключу `cell_types_header in ${sheet_name}` строку вида '0011212'. 
   * Где цифра означает 0 - DEFAULT, 1 - CHECKBOX, 2 - DROPDOWN
   */
  export function updateCellTypeHeaderCache(
    values_coord: {x: number, y: number},
    columns: number,
    sheet: GoogleAppsScript.Spreadsheet.Sheet
  )
  {
    let x = values_coord.x;
    const y = values_coord.y;
    let cache_str = '';
    while(columns)
    {

      x++;
      cache_str +=  AutoColumn.get_cell_type(y, x, sheet);
      columns--;
    }

    CacheService.getScriptCache()?.put(`cell_types_header: ${sheet.getSheetName()}`, cache_str);
  }

  
  export const get_cell_type = (
    y: number,
    x: number, 
    sheet: GoogleAppsScript.Spreadsheet.Sheet)
  : Constants.EnumCellType => 
  {
    const cell = sheet.getRange(y, x);
    const criteria_type = cell.getDataValidation()?.getCriteriaType();

    const { CHECKBOX, VALUE_IN_LIST, VALUE_IN_RANGE } = SpreadsheetApp.DataValidationCriteria;
    switch(criteria_type)
    {
      case CHECKBOX: return Constants.EnumCellType.CHECKBOX;
      case VALUE_IN_LIST: return Constants.EnumCellType.DROPDOWN;
      case VALUE_IN_RANGE: return Constants.EnumCellType.DROPDOWN;
      default: return Constants.EnumCellType.DEFAULT;
    }
    
  };
  
  /**
     * Класс хранящий данные для других классов
     */
    export class DataSource
    {
      
      constructor(
        public header_coords: {y: number, x: number},
        public values_coords: {y: number, x: number},
        public sheet: GoogleAppsScript.Spreadsheet.Sheet
      )
      {        
      }

      get columns()
      {
        return this.sheet.getLastColumn();
      }

      get values_range()
      {
        const { y, x } = this.values_coords;
        
        return this.sheet.getRange(y, x, 1, this.columns);
      }
    }

    /**
     * Класс с механизмами автозаполнения
     */
    export class AutoColumnMechanism
    {
      cache_values = '';

      constructor(private data_source: DataSource)
      {
        this.updateCellTypesHeader();
      }
      

      /**
       * @param x номер координаты
       * Остальные параметры нужны как адаптеры
       */
      getCellType = (y: number, x: number, sheet: GoogleAppsScript.Spreadsheet.Sheet ) => 
      {
        const start_x = this.data_source.values_coords.x;
        
        // Logger.log(this.cache_values[x - start_x]);
        // Logger.log(JSON.stringify(
        //   {
        //     x: x,
        //     x_start_x: x - start_x,
        //     value: this.cache_values[x - start_x],
        //     this_cache_values: this.cache_values
        //   }
        // ));
        
        return Number(this.cache_values[x - start_x - 1]) as Constants.EnumCellType;
      };

      /**
       * Используется для обновление кеша когда открывается таблица 
       * TODO и когда меняются аннотации
       */
      updateCellTypesHeader = () => 
      {
        AutoColumn.updateCellTypeHeaderCache(
          this.data_source.values_coords,
          this.data_source.sheet.getLastColumn(),
          this.data_source.sheet
        );

        const cache = CacheService.getScriptCache();
        const sheet = this.data_source.sheet;

        this.cache_values = cache.get(`cell_types_header: ${sheet.getSheetName()}`) ?? '';
      };
      

      /**
       * @param y - номер строки для автозаполнения
       */
      do = (y: number) => 
      {
        AutoColumn.AutoColumnFunc({ 
          sheet: this.data_source.sheet,
          size: { 
            columns: this.data_source.columns,
            rows: y
          },
          header_coords: this.data_source.header_coords,
          values_coords: this.data_source.values_coords,
        }, this.getCellType);
      };
    }

}


