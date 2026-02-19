import {
  describe,
  expect,
  it
} from 'vitest';

import {
  cellRefA1,
  parseCellRef,
  parseOperation
} from '../src/parsers.ts';

describe('cellRefA1', () => {
  it('converts row 0 col 0 to A1', () => {
    expect(cellRefA1(0, 0)).toBe('A1');
  });

  it('converts row 0 col 4 to E1', () => {
    expect(cellRefA1(0, 4)).toBe('E1');
  });

  it('converts row 3 col 2 to C4', () => {
    expect(cellRefA1(3, 2)).toBe('C4');
  });
});

describe('parseCellRef', () => {
  it('parses A1', () => {
    expect(parseCellRef('A1')).toEqual({ columnId: 1, rowId: 1 });
  });

  it('parses E5', () => {
    expect(parseCellRef('E5')).toEqual({ columnId: 5, rowId: 5 });
  });

  it('is case-insensitive', () => {
    expect(parseCellRef('b3')).toEqual({ columnId: 2, rowId: 3 });
  });

  it('throws for invalid ref', () => {
    expect(() => parseCellRef('ZZ')).toThrow('Bad cell ref');
  });

  it('throws for empty string', () => {
    expect(() => parseCellRef('')).toThrow('Bad cell ref');
  });
});

describe('parseOperation', () => {
  it('parses value assignment =5', () => {
    expect(parseOperation('=5', 1)).toEqual({ type: 'value', value: 5 });
  });

  it('throws for =N with multiple cells', () => {
    expect(() => parseOperation('=3', 2)).toThrow('single cell');
  });

  it('parses candidates 123', () => {
    expect(parseOperation('123', 1)).toEqual({ type: 'candidates', values: [1, 2, 3] });
  });

  it('parses strikethrough -45', () => {
    expect(parseOperation('-45', 1)).toEqual({ type: 'strikethrough', values: [4, 5] });
  });

  it('parses clearance x', () => {
    expect(parseOperation('x', 1)).toEqual({ type: 'clear' });
  });

  it('parses clearance X (uppercase)', () => {
    expect(parseOperation('X', 1)).toEqual({ type: 'clear' });
  });

  it('throws for invalid value after =', () => {
    expect(() => parseOperation('=0', 1)).toThrow('=N expects a single digit 1-9');
  });

  it('throws for invalid strikethrough digits', () => {
    expect(() => parseOperation('-abc', 1)).toThrow('-digits expects digits 1-9');
  });
});
