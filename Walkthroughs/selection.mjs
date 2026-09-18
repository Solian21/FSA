export function pickLowestRandom(
  people,
  areaId,
  amountNeeded,
  deprioritizedIds = new Set()
) {
  const selected = [];

  while (selected.length < amountNeeded) {
    const remaining = people.filter(
      person => !selected.some(selectedPerson => selectedPerson.id === person.id)
    );

    if (remaining.length === 0) {
      break;
    }

    const lowestCount = Math.min(
      ...remaining.map(person => person.counts?.[areaId] ?? 0)
    );

    const tiedPeople = remaining.filter(
      person => (person.counts?.[areaId] ?? 0) === lowestCount
    );

    /*
     * Counts are the primary priority. Within a tie, prefer people who
     * were not in the most recent assignment, then randomize the tie.
     */
    tiedPeople.sort((personA, personB) => {
      const personAWasRecentlyAssigned = deprioritizedIds.has(personA.id);
      const personBWasRecentlyAssigned = deprioritizedIds.has(personB.id);

      if (personAWasRecentlyAssigned !== personBWasRecentlyAssigned) {
        return personAWasRecentlyAssigned ? 1 : -1;
      }

      return Math.random() - 0.5;
    });

    for (const person of tiedPeople) {
      if (selected.length >= amountNeeded) {
        break;
      }

      selected.push(person);
    }
  }

  return selected;
}
