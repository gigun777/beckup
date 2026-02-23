import { journalsForSpace } from './journal_tree_core.js';

export function normalizeLocation({ spaces, journals, lastLoc }) {
  const root = spaces.find((s) => s.parentId === null) ?? spaces[0] ?? null;
  if (!root) return { activeSpaceId: null, activeJournalId: null };

  let activeSpaceId = lastLoc?.activeSpaceId;
  if (!spaces.some((s) => s.id === activeSpaceId)) activeSpaceId = root.id;

  const currentSpaceJournals = journalsForSpace(journals, activeSpaceId);
  let activeJournalId = lastLoc?.activeJournalId;
  if (!currentSpaceJournals.some((j) => j.id === activeJournalId)) {
    const firstLevel = currentSpaceJournals.find((j) => j.parentId === activeSpaceId);
    activeJournalId = firstLevel?.id ?? null;
  }

  return { activeSpaceId, activeJournalId };
}
