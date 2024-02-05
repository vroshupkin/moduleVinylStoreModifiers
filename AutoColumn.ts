namespace AutoColumn
{   
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
    export class DataSourceWithCache
    { 
      cache_values = '';

      constructor(
        public header_coords: {y: number, x: number},
        public values_coords: {y: number, x: number},
        public sheet: GoogleAppsScript.Spreadsheet.Sheet,
      ){}

      /** Количество колонок */
      get columns()
      {
        return this.sheet.getLastColumn();
      }

      get values_row()
      {
        const { y, x } = this.values_coords;
        
        return this.sheet.getRange(y, x, 1, this.columns);
      }

      get header_row()
      {
        const { y, x } = this.header_coords;
        
        return this.sheet.getRange(y, x, 1, this.columns);
      }
      
      
      /**
       * @param x номер координаты
       * Остальные параметры нужны как адаптеры
       */
      getCellType = (x: number) => 
      {
        const start_x = this.values_coords.x;
        
        return Number(this.cache_values[x - start_x - 1]) as Constants.EnumCellType;
      };

      /**
       * Используется для обновление кеша когда открывается таблица
       */
      updateCellTypesHeader = () => 
      {
        AutoColumn.updateCellTypeHeaderCache(
          this.values_coords,
          this.sheet.getLastColumn(),
          this.sheet
        );

        const cache = CacheService.getScriptCache();
        this.cache_values = cache.get(`cell_types_header: ${this.sheet.getSheetName()}`) ?? '';
        Logger.log(this.cache_values);
      };
    }

    /**
     * Класс с механизмами автозаполнения
     */
    export class AutoColumnMechanism
    {
      constructor(private data_source: DataSourceWithCache)
      {
        this.data_source.updateCellTypesHeader();
      }
      
      /**
       * @param y - номер строки для автозаполнения
       */
      do = (y: number) => 
      {
        Logger.log(y);
        const { header_row: header_range, values_coords } = this.data_source;

        const notes = header_range.getNotes()[0];
        Logger.log(notes);
        if(!Array.isArray(notes)) { return; }

        
        // const rows = y - values_coords.y + 1;

        for (const [ i, note ] of notes.entries()) 
        { 
          if(i === 0)
          {
            continue;   
          }

          const x = values_coords.x + i;
          const note_type = (note + '').trim();
          

          // if(note_type.includes('DiscogsFormat'))
          // {
          //   Tools.timestampDecorator(this.discogsFormatAuto, 'discogsFormat')(y, x, this.getParam(note));
          // }
          // Logger.log(note_type);
          if(note_type.toLowerCase().includes('auto'))
          {
            // Logger.log(note_type);
            this.autoFunction(y, x);
          }
          else if(note_type.includes('SetValues'))
          {
            this.setValues(y, x, this.getParam(note));
          }


        }
      };

      private autoFunction = (y: number, x: number) => 
      {
        const { getCellType, sheet, values_coords } = this.data_source;
        const rows = y - values_coords.y + 1;

        const cell_type = getCellType(x);
    
        if(cell_type === Constants.EnumCellType.DEFAULT)
        {
          const pattern_cell = sheet.getRange(values_coords.y, x);
          const insert_column = sheet.getRange(values_coords.y, x, rows);
          
          pattern_cell.autoFill(insert_column, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); 
      
        }
        else if(cell_type === Constants.EnumCellType.CHECKBOX)
        {
          // Logger.log('Constants.EnumCellType.CHECKBOX');
          const insert_cell = sheet.getRange(y, x);
          const check_box_data_validation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      
          insert_cell.setDataValidation(check_box_data_validation);
          insert_cell.setValue(false);
        }
        else if(cell_type === Constants.EnumCellType.DROPDOWN)
        { 
          const insert_cell = sheet.getRange(y, x);         
          const pattern_cell = sheet.getRange(values_coords.y, x); 

          pattern_cell.copyTo(insert_cell);
        }
      };

      private discogsFormatAuto = (y: number, x: number, param: string) =>
      {
        
        // @ts-ignore
        // this.autoColumnFabricFunction(y, x, param, Discogs.formatHandle);
        
      };

      private setValues = (y: number, x: number, param: string) => 
      {
        this.autoColumnFabricFunction(y, x, param, (a: string) => a);
      };

      private autoColumnFabricFunction = (y: number, x: number, param: string, fn: (val: any) => any) =>
      {
        const { sheet, values_coords } = this.data_source;
        
        const source_x = Tools.calcXPosition(param);

        const insert_column = sheet.getRange(values_coords.y, x, y);
        const source_column = sheet.getRange(values_coords.y, source_x, y);

        Tools.fillTableByTable(insert_column, source_column, fn);
      };
    
      /**
       * @example "DiscogsFormat(K)" => "K"
       */
      private getParam = (str: string) =>
      {
        return str.slice(
          str.indexOf('(')+ 1,
          str.indexOf(')')
        );
      };
    }
}


