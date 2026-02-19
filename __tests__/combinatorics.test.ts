import {
  describe,
  expect,
  it
} from 'vitest';

import {
  computeGridBoundaries,
  evaluateTuple,
  generateSubsets
} from '../src/combinatorics.ts';

describe('evaluateTuple', () => {
  it('returns null for empty tuple', () => {
    expect(evaluateTuple([], '+')).toBeNull();
  });

  it('returns the value for single-element tuple', () => {
    expect(evaluateTuple([5], '+')).toBe(5);
  });

  it('evaluates addition', () => {
    expect(evaluateTuple([2, 3], '+')).toBe(5);
    expect(evaluateTuple([1, 2, 3], '+')).toBe(6);
  });

  it('evaluates subtraction (absolute difference)', () => {
    expect(evaluateTuple([5, 3], '-')).toBe(2);
    expect(evaluateTuple([3, 5], '-')).toBe(2);
  });

  it('returns null for subtraction with more than 2 elements', () => {
    expect(evaluateTuple([1, 2, 3], '-')).toBeNull();
  });

  it('evaluates multiplication', () => {
    expect(evaluateTuple([2, 3], 'x')).toBe(6);
    expect(evaluateTuple([2, 3], '*')).toBe(6);
    expect(evaluateTuple([1, 2, 3], 'x')).toBe(6);
  });

  it('evaluates division', () => {
    expect(evaluateTuple([6, 3], '/')).toBe(2);
    expect(evaluateTuple([3, 6], '/')).toBe(2);
  });

  it('returns null for non-integer division', () => {
    expect(evaluateTuple([5, 3], '/')).toBeNull();
  });

  it('returns null for division with more than 2 elements', () => {
    expect(evaluateTuple([1, 2, 3], '/')).toBeNull();
  });

  it('returns null for division by zero', () => {
    expect(evaluateTuple([0, 5], '/')).toBeNull();
  });

  it('returns null for unknown operator', () => {
    expect(evaluateTuple([1, 2], '%')).toBeNull();
  });
});

describe('generateSubsets', () => {
  it('generates empty subset for k=0', () => {
    expect(generateSubsets([1, 2, 3], 0)).toEqual([[]]);
  });

  it('generates all singletons for k=1', () => {
    expect(generateSubsets([1, 2, 3], 1)).toEqual([[1], [2], [3]]);
  });

  it('generates all pairs for k=2', () => {
    expect(generateSubsets([1, 2, 3], 2)).toEqual([[1, 2], [1, 3], [2, 3]]);
  });

  it('generates full set for k=n', () => {
    expect(generateSubsets([1, 2, 3], 3)).toEqual([[1, 2, 3]]);
  });

  it('returns empty array when k > n', () => {
    expect(generateSubsets([1, 2], 3)).toEqual([]);
  });
});

describe('computeGridBoundaries', () => {
  it('computes boundaries for a 2x2 grid with 2 cages', () => {
    const cages = [
      { cells: ['A1', 'A2'], value: 3 },
      { cells: ['B1', 'B2'], value: 3 }
    ];
    const { horizontalBounds, verticalBounds } = computeGridBoundaries(cages, 2);
    // A1=(r0,c0) A2=(r1,c0) are same cage; B1=(r0,c1) B2=(r1,c1) same cage
    // Vertical bound at r=0,c=1: cage of (r0,c0)=0 vs cage of (r0,c1)=1 => true
    expect(verticalBounds[0]).toEqual([true]);
    expect(verticalBounds[1]).toEqual([true]);
    // Horizontal bound at r=1: cage of (r0,c0)=0 vs cage of (r1,c0)=0 => false
    expect(horizontalBounds[0]).toEqual([false, false]);
  });
});
