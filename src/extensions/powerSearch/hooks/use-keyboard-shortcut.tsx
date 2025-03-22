import { useCallback, useEffect } from 'react';

/**
 * Type definition for a keyboard shortcut combination
 * @typedef {Object} KeyCombination
 * @property {string} key - The key to listen for (e.g., 'ArrowUp', 'Enter')
 * @property {boolean} [ctrlKey] - Whether Ctrl key should be pressed
 * @property {boolean} [metaKey] - Whether Meta/Command key should be pressed
 * @property {boolean} [altKey] - Whether Alt key should be pressed
 * @property {boolean} [shiftKey] - Whether Shift key should be pressed
 */
type KeyCombination = {
  key: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  altKey?: boolean;
  shiftKey?: boolean;
};

/**
 * Options for the keyboard shortcut hook
 * @typedef {Object} KeyboardShortcutOptions
 * @property {boolean} [preventDefault=true] - Whether to prevent default browser behavior
 * @property {boolean} [stopPropagation=true] - Whether to stop event propagation
 * @property {boolean} [useCapture=true] - Whether to use capture phase for event listener
 */
type KeyboardShortcutOptions = {
  preventDefault?: boolean;
  stopPropagation?: boolean;
  useCapture?: boolean;
};

/**
 * Custom hook to handle keyboard shortcuts
 *
 * @param {KeyCombination|KeyCombination[]} keyCombination - Single or multiple key combinations to listen for
 * @param {Function} callback - Function to call when the key combination is pressed
 * @param {KeyboardShortcutOptions} [options] - Additional options for the keyboard shortcut
 *
 * @example
 *
 * useKeyboardShortcut({ key: 'k', ctrlKey: true }, () => {
 *   console.log('Ctrl+K pressed');
 * });
 *
 * @example
 *
 * useKeyboardShortcut(
 *   [{ key: 'ArrowUp', ctrlKey: true }, { key: 'ArrowUp', metaKey: true }],
 *   openSearchDialog
 * );
 */
export const useKeyboardShortcut = (
  keyCombination: KeyCombination | KeyCombination[],
  callback: () => void,
  options: KeyboardShortcutOptions = {}
) => {
  const {
    preventDefault = true,
    stopPropagation = true,
    useCapture = true,
  } = options;

  const combinations = Array.isArray(keyCombination)
    ? keyCombination
    : [keyCombination];

  /**
   * Event handler for keydown events
   * @param {KeyboardEvent} event - The keyboard event
   * @returns {boolean|undefined} - Returns false for older browsers to prevent default
   */
  const handleKeyDown = useCallback(
    (event: KeyboardEvent) => {
      if (
        event.target instanceof HTMLElement &&
        (event.target.tagName === 'INPUT' ||
          event.target.tagName === 'TEXTAREA' ||
          event.target.isContentEditable)
      ) {
        return;
      }

      const matchesCombination = combinations.some((combo) => {
        const keyMatches = event.key.toLowerCase() === combo.key.toLowerCase();
        const ctrlMatches =
          combo.ctrlKey === undefined || event.ctrlKey === combo.ctrlKey;
        const metaMatches =
          combo.metaKey === undefined || event.metaKey === combo.metaKey;
        const altMatches =
          combo.altKey === undefined || event.altKey === combo.altKey;
        const shiftMatches =
          combo.shiftKey === undefined || event.shiftKey === combo.shiftKey;

        return (
          keyMatches && ctrlMatches && metaMatches && altMatches && shiftMatches
        );
      });

      if (matchesCombination) {
        if (preventDefault) event.preventDefault();
        if (stopPropagation) event.stopPropagation();

        callback();

        return false;
      }
    },
    [combinations, callback, preventDefault, stopPropagation]
  );

  useEffect(() => {
    // Add the keydown listener
    const addKeydownListener = () => {
      document.addEventListener('keydown', handleKeyDown, useCapture);
    };

    // Remove the keydown listener
    const removeKeydownListener = () => {
      document.removeEventListener('keydown', handleKeyDown, useCapture);
    };

    // Handle window focus event - reattach the listener when window regains focus
    const handleWindowFocus = () => {
      // Remove first to avoid duplicate listeners
      removeKeydownListener();
      addKeydownListener();
    };

    // Add initial listeners
    addKeydownListener();
    window.addEventListener('focus', handleWindowFocus);

    // Cleanup all listeners on unmount
    return () => {
      removeKeydownListener();
      window.removeEventListener('focus', handleWindowFocus);
    };
  }, [handleKeyDown, useCapture]);
};
