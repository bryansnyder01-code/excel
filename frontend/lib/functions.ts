// Excel functions offered by formula autocomplete, with their arguments and a one-line description.

export interface FunctionInfo {
  name: string;
  args: string[];
  desc: string;
}

const F = (name: string, args: string, desc: string): FunctionInfo => ({ name, args: args ? args.split(", ") : [], desc });

export const FUNCTIONS: FunctionInfo[] = [
  F("SUM", "number1, [number2], ...", "Adds numbers or ranges"),
  F("AI", "prompt, [input1], [input2], ...", "Asks AI about the inputs, e.g. categorise or translate them"),
  F("AVERAGE", "number1, [number2], ...", "Average of the numbers"),
  F("COUNT", "value1, [value2], ...", "Counts cells that contain numbers"),
  F("COUNTA", "value1, [value2], ...", "Counts cells that aren't empty"),
  F("COUNTBLANK", "range", "Counts empty cells"),
  F("MIN", "number1, [number2], ...", "Smallest number"),
  F("MAX", "number1, [number2], ...", "Largest number"),
  F("MEDIAN", "number1, [number2], ...", "Middle number"),
  F("ROUND", "number, num_digits", "Rounds to a number of digits"),
  F("ROUNDUP", "number, num_digits", "Rounds away from zero"),
  F("ROUNDDOWN", "number, num_digits", "Rounds towards zero"),
  F("ABS", "number", "Absolute value"),
  F("IF", "logical_test, value_if_true, [value_if_false]", "One value if a condition is true, another if not"),
  F("IFS", "test1, value1, [test2, value2], ...", "First value whose condition is true"),
  F("IFERROR", "value, value_if_error", "A fallback when a formula gives an error"),
  F("IFNA", "value, value_if_na", "A fallback for #N/A"),
  F("AND", "logical1, [logical2], ...", "TRUE if every condition is true"),
  F("OR", "logical1, [logical2], ...", "TRUE if any condition is true"),
  F("NOT", "logical", "Flips TRUE and FALSE"),
  F("SUMIF", "range, criteria, [sum_range]", "Adds cells that meet one condition"),
  F("SUMIFS", "sum_range, criteria_range1, criteria1, ...", "Adds cells that meet several conditions"),
  F("COUNTIF", "range, criteria", "Counts cells that meet one condition"),
  F("COUNTIFS", "criteria_range1, criteria1, ...", "Counts cells that meet several conditions"),
  F("AVERAGEIF", "range, criteria, [average_range]", "Average of cells that meet one condition"),
  F("AVERAGEIFS", "average_range, criteria_range1, criteria1, ...", "Average of cells that meet several conditions"),
  F("MAXIFS", "max_range, criteria_range1, criteria1, ...", "Largest value that meets conditions"),
  F("MINIFS", "min_range, criteria_range1, criteria1, ...", "Smallest value that meets conditions"),
  F("XLOOKUP", "lookup_value, lookup_array, return_array, [if_not_found]", "Finds a value and returns the matching item"),
  F("VLOOKUP", "lookup_value, table_array, col_index_num, [range_lookup]", "Looks up a value in the first column of a table"),
  F("HLOOKUP", "lookup_value, table_array, row_index_num, [range_lookup]", "Looks up a value in the first row of a table"),
  F("INDEX", "array, row_num, [column_num]", "The value at a position in a range"),
  F("MATCH", "lookup_value, lookup_array, [match_type]", "The position of a value in a range"),
  F("CHOOSE", "index_num, value1, [value2], ...", "Picks a value from a list by number"),
  F("SUMPRODUCT", "array1, [array2], ...", "Multiplies ranges and adds the results"),
  F("PRODUCT", "number1, [number2], ...", "Multiplies numbers"),
  F("POWER", "number, power", "A number raised to a power"),
  F("SQRT", "number", "Square root"),
  F("MOD", "number, divisor", "Remainder after division"),
  F("INT", "number", "Rounds down to a whole number"),
  F("CEILING", "number, significance", "Rounds up to a multiple"),
  F("FLOOR", "number, significance", "Rounds down to a multiple"),
  F("RANK", "number, ref, [order]", "Rank of a number in a list"),
  F("LARGE", "array, k", "The k-th largest value"),
  F("SMALL", "array, k", "The k-th smallest value"),
  F("STDEV", "number1, [number2], ...", "Standard deviation of a sample"),
  F("VAR", "number1, [number2], ...", "Variance of a sample"),
  F("PERCENTILE", "array, k", "The k-th percentile"),
  F("PMT", "rate, nper, pv, [fv], [type]", "Loan payment per period"),
  F("IPMT", "rate, per, nper, pv, [fv], [type]", "Interest part of a loan payment"),
  F("PPMT", "rate, per, nper, pv, [fv], [type]", "Principal part of a loan payment"),
  F("FV", "rate, nper, pmt, [pv], [type]", "Future value of an investment"),
  F("PV", "rate, nper, pmt, [fv], [type]", "Present value of an investment"),
  F("NPV", "rate, value1, [value2], ...", "Net present value"),
  F("IRR", "values, [guess]", "Internal rate of return"),
  F("RATE", "nper, pmt, pv, [fv], [type], [guess]", "Interest rate per period"),
  F("NPER", "rate, pmt, pv, [fv], [type]", "Number of payment periods"),
  F("TODAY", "", "Today's date"),
  F("NOW", "", "The current date and time"),
  F("DATE", "year, month, day", "Builds a date"),
  F("YEAR", "date", "The year of a date"),
  F("MONTH", "date", "The month of a date, 1 to 12"),
  F("DAY", "date", "The day of the month"),
  F("WEEKDAY", "date, [return_type]", "Day of the week as a number"),
  F("EOMONTH", "start_date, months", "Last day of a month"),
  F("EDATE", "start_date, months", "The same day some months later"),
  F("DATEDIF", "start_date, end_date, unit", "Years, months or days between dates"),
  F("NETWORKDAYS", "start_date, end_date, [holidays]", "Working days between dates"),
  F("WORKDAY", "start_date, days, [holidays]", "A date some working days away"),
  F("CONCAT", "text1, [text2], ...", "Joins text"),
  F("TEXTJOIN", "delimiter, ignore_empty, text1, ...", "Joins text with a separator"),
  F("LEFT", "text, [num_chars]", "Characters from the start"),
  F("RIGHT", "text, [num_chars]", "Characters from the end"),
  F("MID", "text, start_num, num_chars", "Characters from the middle"),
  F("LEN", "text", "Number of characters"),
  F("TRIM", "text", "Removes extra spaces"),
  F("UPPER", "text", "Converts to capitals"),
  F("LOWER", "text", "Converts to lowercase"),
  F("PROPER", "text", "Capitalises each word"),
  F("SUBSTITUTE", "text, old_text, new_text, [instance_num]", "Replaces text"),
  F("FIND", "find_text, within_text, [start_num]", "Position of text (case-sensitive)"),
  F("SEARCH", "find_text, within_text, [start_num]", "Position of text (any case)"),
  F("TEXT", "value, format_text", "Formats a number as text"),
  F("VALUE", "text", "Converts text to a number"),
  F("ISBLANK", "value", "TRUE if the cell is empty"),
  F("ISNUMBER", "value", "TRUE if the value is a number"),
  F("ISTEXT", "value", "TRUE if the value is text"),
  F("ISERROR", "value", "TRUE if the value is an error"),
  F("UNIQUE", "array", "Distinct values (Excel 365)"),
  F("FILTER", "array, include, [if_empty]", "Rows that meet a condition (Excel 365)"),
  F("SORT", "array, [sort_index], [sort_order]", "Sorted range (Excel 365)"),
];

const BY_NAME = new Map(FUNCTIONS.map((f) => [f.name, f]));

export const functionInfo = (name: string) => BY_NAME.get(name.toUpperCase());

export interface FormulaContext {
  /** The function name being typed at the cursor, if any. */
  partial: { text: string; start: number } | null;
  /** The function call the cursor is inside, and which argument it's on. */
  call: { name: string; arg: number } | null;
}

/** Work out what the user is typing at `caret` in a formula. */
export function formulaContext(value: string, caret: number): FormulaContext {
  if (!value.startsWith("=")) return { partial: null, call: null };
  const before = value.slice(0, caret);
  const stack: { name: string; arg: number }[] = [];
  let inString = false;
  let word = "";
  for (let i = 1; i < before.length; i++) {
    const ch = before[i];
    if (ch === '"') {
      inString = !inString;
      word = "";
      continue;
    }
    if (inString) continue;
    if (/[A-Za-z0-9_.]/.test(ch)) {
      word += ch;
      continue;
    }
    if (ch === "(") stack.push({ name: word.toUpperCase(), arg: 0 });
    else if (ch === ")") stack.pop();
    else if (ch === "," && stack.length) stack[stack.length - 1].arg++;
    word = "";
  }
  const m = inString ? null : /(^|[^A-Za-z0-9_.$!'"])([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
  // A run of letters followed by digits is a cell reference (B12), not a function name.
  const partial = m && !/^[A-Za-z]{1,3}\d+$/.test(m[2]) ? { text: m[2], start: caret - m[2].length } : null;
  const call = [...stack].reverse().find((c) => functionInfo(c.name)) ?? null;
  return { partial, call };
}

export function suggestFunctions(partial: string, limit = 8): FunctionInfo[] {
  const p = partial.toUpperCase();
  return FUNCTIONS.filter((f) => f.name.startsWith(p)).slice(0, limit);
}
