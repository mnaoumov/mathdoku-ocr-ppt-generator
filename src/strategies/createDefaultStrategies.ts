import type { Strategy } from './Strategy.ts';

import { HiddenSingleStrategy } from './HiddenSingleStrategy.ts';
import { NakedSetStrategy } from './NakedSetStrategy.ts';
import { SingleCandidateStrategy } from './SingleCandidateStrategy.ts';

const MIN_NAKED_SET_SIZE = 2;

export function createDefaultStrategies(size: number): Strategy[] {
  return [
    new SingleCandidateStrategy(),
    new HiddenSingleStrategy(),
    ...Array.from({ length: size - MIN_NAKED_SET_SIZE }, (_, i) => new NakedSetStrategy(i + MIN_NAKED_SET_SIZE))
  ];
}
