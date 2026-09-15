export function ok(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(`consumer fixture: ${message}`);
}
