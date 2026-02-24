import { describe, it, expect } from 'vitest';
import { emuToInches, emuAngleToDegrees, hundredthsToPoints, emuToPoints } from '../src/mapper/units.js';

describe('emuToInches', () => {
  it('converts 914400 EMU to 1 inch', () => {
    expect(emuToInches(914400)).toBe(1);
  });

  it('converts 457200 EMU to 0.5 inches', () => {
    expect(emuToInches(457200)).toBe(0.5);
  });

  it('converts 0 EMU to 0 inches', () => {
    expect(emuToInches(0)).toBe(0);
  });
});

describe('emuAngleToDegrees', () => {
  it('converts 5400000 to 90 degrees', () => {
    expect(emuAngleToDegrees(5400000)).toBe(90);
  });

  it('converts 0 to 0 degrees', () => {
    expect(emuAngleToDegrees(0)).toBe(0);
  });
});

describe('hundredthsToPoints', () => {
  it('converts 3200 hundredths to 32 points', () => {
    expect(hundredthsToPoints(3200)).toBe(32);
  });

  it('converts 1000 hundredths to 10 points', () => {
    expect(hundredthsToPoints(1000)).toBe(10);
  });
});

describe('emuToPoints', () => {
  it('converts 12700 EMU to 1 point', () => {
    expect(emuToPoints(12700)).toBe(1);
  });

  it('converts 25400 EMU to 2 points', () => {
    expect(emuToPoints(25400)).toBe(2);
  });
});
