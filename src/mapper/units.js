// EMU -> inches (4 decimal places)
export function emuToInches(emu) {
  return Math.round((emu / 914400) * 10000) / 10000;
}

// EMU angle -> degrees
export function emuAngleToDegrees(emuAngle) {
  return emuAngle / 60000;
}

// Hundredths of a point -> points
export function hundredthsToPoints(val) {
  return val / 100;
}

// EMU -> points (for line widths)
export function emuToPoints(emu) {
  return Math.round((emu / 12700) * 100) / 100;
}
