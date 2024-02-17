

namespace Tools
{
  
  export function getSheet(name: Constants.TSheetsNames)
  {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  }
  /**
   * Используется для взаимодействия с дропдаун ячейкой
   */
  export class Dropdown
  {
    private values: string[];

    /**
     * @param cell ячейка с одним значением 
     */
    constructor(cell: GoogleAppsScript.Spreadsheet.Range)
    {
      // TODO Сделать поверку что DataValidations содерржит dropdown range
      const dropdown_range = cell.getDataValidations()?.[0][0]
        ?.getCriteriaValues()[0] as GoogleAppsScript.Spreadsheet.Range;
      
      if(!dropdown_range)
      {
        throw new UiError('');
      }

      this.values = dropdown_range.getValues().flat();
    }

    includes = (str: string) => 
    {
      return this.values.includes(str);
    };
  }

  interface IToString {
    toString: () => string
  }
  /**
   * Ошибка с всплывающим окном
   */
  export class UiError extends Error
  {
    constructor(message: IToString | null | undefined) 
    {
      super();

      this.name = 'UiError';
      this.message = message + '';
      SpreadsheetApp.getUi().alert(this.stack + '');
    }
  }


  /**
   * Переводит индекс столбца в буквенно обозначение
   * @param num индекс столбца
   * @example ConverToChar(1) = "A"; ConverToChar(10) = "J"; ConverToChar(27) = "AA"
   */
  export const ConverToChar = (num: number) => 
  {
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

    const latters: string[] = [];
    do
    {
      num -= 1;
      latters.push(alphabet[num % 26]);
      num = Math.floor(num / 26);
    }while(num > 0);

    let output = '';
    while(latters.length)
    {
      output += latters.pop();
    }
    
    return output;
  };

  /**
   * Создает строку из len пробелов
   * @example generate_n_spaces(5) => '     '
   */
  export const generate_n_spaces = (len: number): string =>
  {
    let out = '';
    while(len--) {out += ' ';}
    
    return out;
  }; 

  /**
     * Генерирует название колонок в JSON строке для использования в коде. Для использования нужно выделить строку с 
     * нужными строками 
     * @param range Первая ячейка строки
     * @param offset 
     */
  export const GenerateColumnIndexes = (range: GoogleAppsScript.Spreadsheet.Range, offset = 1) => 
  {
    // Перевести в выделение строки
    const column_names = range.getValues()[0];
    const have_dict: {[s: string]: any} = {};
    column_names.forEach((key, i) => 
    {
      key = key + '';
      while(have_dict[key] != undefined)
      {
        key += 'COPY';
      }

      have_dict[key] = Tools.ConverToChar(i + offset);
    });

    new Tools.UiError(JSON.stringify(have_dict));

    let output = '';
    let count = 0;
    for (const key in have_dict) 
    {
      const val = have_dict[key];
      
      // Поиск символа новой строки и замена на валидный символ
      const ind_newline = val.indexOf('\n');
      if(ind_newline >= 0)
      {
        val[ind_newline] = '\\\n';
      }

      const new_line = count == 2? '\n': '';
      const columns_string = `'${key}': ${val}, ${new_line}`;
      
      count = count == 2? 0: ++count;
      output += columns_string;
    }

    return `{${output}}`;
  };

  // /**
  //  * Автоматическое протягивание
  //  */
  // export const auto_column = 
  // (sheet: GoogleAppsScript.Spreadsheet.Sheet) => (start_y: number, start_x: number, rows_num: number) => 
  // {
  //   const start_cell = sheet.getRange(start_y, start_x);
  //   const destination = sheet.getRange(start_y, start_x, rows_num);
  //   start_cell.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); 
  // };

  
  // /** 
  //  * Ищет в аннотациях названия 'auto' и делает автозаполнение колонки
  //  * @param header_range Строка с названием таблиц и аннотацией
  //  * @param insert_y Строка с формулами изменением данных
  //  * @param headers_cells Строка с заполненными формулами
  //  */
  // export const setAutoColumn = (
  //   params: {
  //     header_range: GoogleAppsScript.Spreadsheet.Range,
  //     header_cells: GoogleAppsScript.Spreadsheet.Range,
  //     insert_y: number,
  //     sheet_name: GoogleAppsScript.Spreadsheet.Sheet
  //   }
  // ) => 
  // { 
  //   const { header_range, header_cells, insert_y, sheet_name: sheet } = params;
     
  //   const notes = header_range.getNotes()[0];

  //   const header_y = header_range.getRow();

  //   for (let i = 0; i < notes.length; i++) 
  //   {      
      
  //     let note = notes[i];
  //     note = note.slice(0, 4);

  //     if(note != 'auto') {continue;}
      
  //     const x =  header_range.getColumn() + i;      

  //     const source_cell = sheet.getRange(header_y + 1, x);
  //     const insert_cell = sheet.getRange(insert_y, x);

  //     const cell_validation_type = Tools.readHeaderCache(i, Constants.CACHE_NAMES.EconomyHeader + '');
  //     if(cell_validation_type === undefined) 
  //     {  
  //       Tools.updateHeaderCache(header_cells, Constants.CACHE_NAMES.EconomyHeader);
  //     }

  //     const { CHECKBOX, DEFAULT, DROPDOWN } = EnumCellType;
      
  //     switch(cell_validation_type + '')
  //     {
  //       case DEFAULT + '': {
  //         Tools.auto_column(sheet)(header_y + 1, x, insert_y - header_y);
  //         break;
  //       }
  //       case DROPDOWN + '': {
  //         const default_dropdown_val = source_cell.getValue();
  //         insert_cell.setValue(default_dropdown_val);
  //         break;  
  //       }
  //       case CHECKBOX + '':{
  //         insert_cell.setValue('');
  //         break;  
  //       }
  //     }
  //   }
    

  // };


  /**
   * Смотрит какого типа ячейка: обычная, чекбокс, dropdown
   * @param cell 
   * @returns 
   */
  export const get_cell_type = (cell: GoogleAppsScript.Spreadsheet.Range): Constants.EnumCellType => 
  {
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
  
  
  export function readHeaderCache(i: number, cache_key: string)
  {
    let arr_str = CacheService.getDocumentCache()?.get(cache_key);
    arr_str = arr_str ?? '';
    const arr = (Array.from<unknown>(arr_str)[i]);

    return arr;
  }
  /**
   * Считает позицию в гугл таблицах
   * A = 1; AA = 27; AB = 28; BA = 53
   */
  export function calcXPosition(str: string)
  {
    str = str.toLocaleLowerCase();
    if(str.length === 1)
    {
      return str[0].charCodeAt(0) - 96;
    }
    else if(str.length === 2)
    {
      const res = str[1].charCodeAt(0) - 97;
      console.log(res);
      
      return res + 27 * (str[0].charCodeAt(0) - 96);
    }
    else
    {
      throw Error('Данная функция считает только позиции с одной или двумя буквами');
    }
  }
  
  
  export class TimestampTools
  {
    timestamps: Date[];
    timestamps_names: string[];
    constructor()

    {
      [ this.timestamps, this.timestamps_names ] = [ [], [] ];
      
    }

    /**
     * 
     * @param name 
     * @example timestamp.tag('a') \n timestamp.tag('a') \n timestamp + '' => 
     * 
     */
    tag(name: string)
    {
      this.timestamps.push(new Date()) && this.timestamps_names.push(name ?? '');
    }

    toString()
    {
      const get_time = (t: Date) => 
      {
        const time = t.getTime() - this.timestamps[0].getTime();
        if(time < 1000)
        {
          return time + 'ms';
        }
        else
        {
          return time / 1000 + 's';
        }

      };
      const timestamps_arr = this.timestamps.map(get_time);
      const timestamps_msg = timestamps_arr.map((_, i) => `[ ${this.timestamps_names[i]}, ${timestamps_arr[i]} ]`);
      
      return timestamps_msg + '';
    
    }

    script_log()
    {
      Logger.log(this.toString());
    }
  }
 

  /**
   * Собирает query search параметр
   * @example url_builder('https://api.moysklad.ru/api/remap/1.2','hello', {a: '10', b: '20'}) => 
   * "https://api.moysklad.ru/api/remap/1.2/hello?a=10&b=20" 
   */
  export const url_builder = (base_url: string, end_point: string, query_params: Record<string, string>) => 
  {
    let query = Object.entries(query_params)
      .map(v => `&${v[0]}=${v[1]}`)
      .reduce((a, b) => a + b);
  
    query = query.slice(1, query.length);

    return `${base_url}/${end_point}?${query}`;

  };

    /** @example getColumn('Таблица_1, 'A') - Заберет все данных из колонки 'A' и вернет их в виде массива */
  export function getColumn(sheet: GoogleAppsScript.Spreadsheet.Sheet, ch: string)
  {       
    const out: any = [];
    
    for (const row of sheet.getRange(`${ch}:${ch}`).getValues()) 
    {
      out.push(row[0]);
    }   

    return out;
  }

  /**
   * Логгер декаратор для функции
   * @param fn Функция которую нужно сдекарировать
   * @param function_name Имя функции, которое будет логироваться
   */
  export function timestampDecorator(fn: unknown, function_name = '')
  {
    if(!(typeof fn === 'function'))
    {
      throw Error('Need to pass function');
    }
    
    const loggerDecoratorOutputFunction = (...args: unknown[]) => 
    {
      const d1 = new Date();
      
      const result = fn(...args);
      const d2 = new Date();
      const sec = (Number(d2) - Number(d1)) / 1000;

      Logger.log(`Функция ${function_name}() отработала за ${sec} секунд`);
      
      return result;      
    }; 

    return loggerDecoratorOutputFunction;
  }

  
  export function is_object(obj: any): obj is {[str: PropertyKey]: unknown}
  {
    return typeof(obj) === 'object' && obj !== null;
  }

  export const is_string = (obj: any): obj is string  => typeof(obj) === 'string';

  export function object_has_keys<T extends string | number>(obj: any, ...keys: T[]): obj is {[k in T] : unknown}
  {
    if(!is_object(obj))
    {
      return false;
    }

    for (const key of keys) 
    {
      if(!(key in obj))
      {
        return false;
      }
    }

    return true;
  } 

  export function is_key_in_object<T>(key: PropertyKey, obj: T): key is keyof T 
  {
    return Object.prototype.hasOwnProperty.call(obj, key);
  }


  /**
   * Берет значения из таблицы source_table. Обрабатывает их обработчиком handler. 
   * И записывает в таблицу insert_table
   */
  export function fillTableByTable<INPUT, OUTPUT>(
    insert_table: GoogleAppsScript.Spreadsheet.Range,
    source_table: GoogleAppsScript.Spreadsheet.Range,
    handler: (val: INPUT) => OUTPUT
  )
  {
    const values = source_table.getValues();
    const res = values.map(arr => arr.map(handler));
    
    insert_table.setValues(res);
  }


  /**
   * В столбце ищет координату строки последней пустой ячейки
   */
  export const getLastRowIndex = (sheet: GoogleAppsScript.Spreadsheet.Sheet, a1notation: string) => 
  {
    const range = sheet.getRange(a1notation);
    const values = range.getValues();
    
    Logger.log(`rows: ${values.length}, columns: ${values[0].length}`);
    
    
    for(let y = 0; y < values.length; y++)
    {
    
      if((values[y][0] + '').trim() === '')
      {
        return y + range.getRow() - 1;
      }
    }

    return -1;
  };
}
