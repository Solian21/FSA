import assert from "node:assert/strict";
import test from "node:test";
import { pickLowestRandom } from "./selection.mjs";

function person(id, count) {
  return { id, counts: { 1: count } };
}

test("a zero-count person beats someone with a higher count", () => {
  const selected = pickLowestRandom(
    [person(1, 0), person(2, 3)],
    1,
    1,
    new Set([1])
  );

  assert.deepEqual(selected.map(item => item.id), [1]);
});

test("a person from the previous assignment is deprioritized on a tie", () => {
  const selected = pickLowestRandom(
    [person(1, 0), person(2, 0)],
    1,
    1,
    new Set([1])
  );

  assert.deepEqual(selected.map(item => item.id), [2]);
});

test("the lowest counts are selected when multiple people are needed", () => {
  const selected = pickLowestRandom(
    [person(1, 2), person(2, 0), person(3, 1)],
    1,
    2
  );

  assert.deepEqual(
    selected.map(item => item.id).sort((a, b) => a - b),
    [2, 3]
  );
});

test("missing counts are treated as zero", () => {
  const selected = pickLowestRandom(
    [{ id: 1, counts: {} }, person(2, 1)],
    1,
    1
  );

  assert.deepEqual(selected.map(item => item.id), [1]);
});
