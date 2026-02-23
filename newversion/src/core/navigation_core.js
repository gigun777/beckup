import { journalsForSpace } from './journal_tree_core.js';

export function canGoBackSpace(activeSpace) {
  return Boolean(activeSpace?.parentId);
}

export function canGoBackJournal(activeJournal, activeSpaceId) {
  if (!activeJournal) return false;
  return activeJournal.parentId !== activeSpaceId;
}

export function currentJournalLabel(state) {
  const items = journalsForSpace(state.journals, state.activeSpaceId);
  if (items.length === 0) return 'Додай журнал';
  return state.activeJournalId
    ? items.find((j) => j.id === state.activeJournalId)?.title ?? 'Журнал'
    : 'Журнал';
}

export function computeNumbering(pathTitles) {
  return pathTitles.map((_, idx) => pathTitles.slice(0, idx + 1).join('.'));
}

export function pushHistory(history, record) {
  return [...history, record].slice(-100);
}

export function plusCreatesOnlyChildren() {
  return true;
}
