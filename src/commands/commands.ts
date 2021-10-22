import setAvailableAction from "./setAvailableAction";
import bookMeetingAction from "./bookMeetingAction";

/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}
const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.setAvailableAction = setAvailableAction;
g.bookMeetingAction = bookMeetingAction;
