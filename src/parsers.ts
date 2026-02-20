import { ensureNonNullable } from './typeGuards.ts';

export type CellOperation = CandidatesOperation | ClearanceOperation | StrikethroughOperation | ValueOperation;

export interface CellRef {
  readonly columnId: number;
  readonly rowId: number;
}

interface CandidatesOperation {
  readonly type: 'candidates';
  readonly values: readonly number[];
}

interface ClearanceOperation {
  readonly type: 'clear';
}

interface StrikethroughOperation {
  readonly type: 'strikethrough';
  readonly values: readonly number[];
}

interface ValueOperation {
  readonly type: 'value';
  readonly value: number;
}

const CHAR_CODE_A = 65;

export function getCellRef(rowId: number, columnId: number): string {
  return String.fromCharCode(CHAR_CODE_A + columnId - 1) + String(rowId);
}

export function parseCellRef(token: string): CellRef {
  const m = /^(?<col>[A-Z])(?<row>[1-9]\d*)$/.exec(token.trim().toUpperCase());
  if (!m) {
    throw new Error(`Bad cell ref: ${token}`);
  }
  const groups = ensureNonNullable(m.groups);
  return {
    columnId: ensureNonNullable(groups['col']).charCodeAt(0) - CHAR_CODE_A + 1,
    rowId: parseInt(ensureNonNullable(groups['row']), 10)
  };
}

export function parseOperation(text: string, cellCount: number): CellOperation {
  if (text.toLowerCase() === 'x') {
    return { type: 'clear' };
  }

  if (text.startsWith('=')) {
    const valueStr = text.substring(1);
    if (!/^[1-9]$/.test(valueStr)) {
      throw new Error('=N expects a single digit 1-9');
    }
    if (cellCount > 1) {
      throw new Error('=N can only be used with a single cell');
    }
    return { type: 'value', value: parseInt(valueStr, 10) };
  }

  if (text.startsWith('-')) {
    const valuesStr = text.substring(1);
    if (!/^[1-9]+$/.test(valuesStr)) {
      throw new Error('-digits expects digits 1-9');
    }
    const values = Array.from(valuesStr, (ch) => parseInt(ch, 10));
    return { type: 'strikethrough', values };
  }

  if (!/^[1-9]+$/.test(text)) {
    throw new Error(`Expected digits 1-9: ${text}`);
  }
  const values = Array.from(text, (ch) => parseInt(ch, 10));
  return { type: 'candidates', values };
}
