/**
 * Tiny shared app state. Enough for Phase 1; may be replaced with a
 * proper observable store in later phases if it grows.
 */

export const appState = {
  currentUser: null,          // { name, email, rpeq, role } — for Phase 1, placeholder
  currentProjectId: null,     // set when a project context is active
  currentInspectionId: null,  // set when an inspection is in progress

  setUser(user)             { this.currentUser = user; },
  setCurrentProject(id)     { this.currentProjectId = id; },
  setCurrentInspection(id)  { this.currentInspectionId = id; }
};
