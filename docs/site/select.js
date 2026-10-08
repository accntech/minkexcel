// Select-only comboboxes with native popover dismissal and a native-select fallback.
import { numericText } from './numbers.mjs';
export function enhanceSelects(root) {
  if (typeof HTMLElement.prototype.showPopover !== 'function') return;
  for (const native of root.querySelectorAll('select')) {
    const field = native.closest('.select-field');
    const name = native.closest('label').firstChild.textContent.trim();
    const trigger = document.createElement('button');
    trigger.type = 'button';
    trigger.className = 'select-trigger';
    trigger.setAttribute('role', 'combobox');
    trigger.setAttribute('aria-label', name);
    trigger.setAttribute('aria-haspopup', 'listbox');
    trigger.setAttribute('aria-expanded', 'false');
    const value = document.createElement('span');
    trigger.append(value, field.querySelector(':scope > .icon').cloneNode(true));
    const popup = document.createElement('div');
    popup.id = `${native.id}-options`;
    popup.className = 'select-popup';
    popup.setAttribute('popover', 'auto');
    popup.setAttribute('role', 'listbox');
    popup.setAttribute('aria-label', name);
    trigger.setAttribute('aria-controls', popup.id);
    const options = [...native.options];
    const items = options.map((option, index) => {
      const item = document.createElement('div');
      item.id = `${native.id}-option-${index}`;
      item.className = 'select-option';
      item.setAttribute('role', 'option');
      const label = document.createElement('span');
      label.innerHTML = numericText(option.textContent);
      item.append(label);
      item.append(field.querySelector('template').content.cloneNode(true));
      popup.append(item);
      return item;
    });
    let active = native.selectedIndex;
    let query = '', lastTyped = 0;
    const isOpen = () => popup.matches(':popover-open');
    const sync = () => {
      value.innerHTML = numericText(options[native.selectedIndex].textContent);
      items.forEach((item, i) => item.setAttribute('aria-selected', String(i === native.selectedIndex)));
    };
    const highlight = index => {
      active = Math.max(0, Math.min(index, items.length - 1));
      items.forEach((item, i) => item.classList.toggle('is-active', i === active));
      trigger.setAttribute('aria-activedescendant', items[active].id);
      items[active].scrollIntoView({ block: 'nearest', behavior: 'instant' });
    };
    const position = () => {
      const box = trigger.getBoundingClientRect();
      const viewport = window.visualViewport;
      const left = viewport?.offsetLeft || 0, top = viewport?.offsetTop || 0;
      const width = viewport?.width || innerWidth, height = viewport?.height || innerHeight;
      const below = top + height - box.bottom - 12, above = box.top - top - 12;
      popup.style.width = `${Math.min(box.width, width - 16)}px`;
      popup.style.maxHeight = `${Math.max(44, Math.max(below, above))}px`;
      const menu = popup.getBoundingClientRect();
      const y = below >= menu.height || below >= above ? box.bottom + 4 : box.top - menu.height - 4;
      popup.style.left = `${Math.max(left + 8, Math.min(box.left, left + width - menu.width - 8))}px`;
      popup.style.top = `${Math.max(top + 8, Math.min(y, top + height - menu.height - 8))}px`;
    };
    const reset = () => {
      trigger.setAttribute('aria-expanded', 'false');
      trigger.removeAttribute('aria-activedescendant');
      query = '';
    };
    const close = () => {
      if (isOpen()) popup.hidePopover();
      reset();
    };
    const commit = () => {
      const changed = native.selectedIndex !== active;
      native.selectedIndex = active;
      sync();
      if (changed) native.dispatchEvent(new Event('change', { bubbles: true }));
      close();
    };
    const open = () => {
      if (isOpen()) return;
      sync();
      popup.showPopover();
      position();
      trigger.setAttribute('aria-expanded', 'true');
      highlight(native.selectedIndex);
    };
    trigger.addEventListener('click', () => isOpen() ? close() : open());
    trigger.addEventListener('keydown', event => {
      const { key } = event;
      if (key === 'Tab') { if (isOpen()) commit(); return; }
      if (key === 'Escape') {
        if (isOpen()) { event.preventDefault(); event.stopPropagation(); close(); }
        return;
      }
      if (key === 'Enter' || key === ' ') {
        event.preventDefault();
        isOpen() ? commit() : open();
        return;
      }
      if (['ArrowDown', 'ArrowUp', 'Home', 'End', 'PageUp', 'PageDown'].includes(key)) {
        event.preventDefault();
        const wasOpen = isOpen();
        open();
        if (key === 'Home') highlight(0);
        else if (key === 'End') highlight(items.length - 1);
        else if (key === 'ArrowUp' && event.altKey) commit();
        else if (wasOpen) highlight(active + (key === 'ArrowDown' ? 1 : key === 'PageDown' ? 10 : key === 'PageUp' ? -10 : -1));
        return;
      }
      if (key.length === 1 && !event.metaKey && !event.ctrlKey && !event.altKey) {
        event.preventDefault();
        open();
        const now = performance.now();
        query = now - lastTyped > 700 ? key.toLowerCase() : query + key.toLowerCase();
        lastTyped = now;
        const prefix = [...query].every(char => char === query[0]) ? query[0] : query;
        const start = prefix.length === 1 ? active + 1 : active;
        for (let i = 0; i < options.length; i++) {
          const index = (start + i) % options.length;
          if (options[index].textContent.toLowerCase().startsWith(prefix)) { highlight(index); break; }
        }
      }
    });
    trigger.addEventListener('blur', () => { if (isOpen()) commit(); });
    popup.addEventListener('pointerdown', event => event.preventDefault());
    items.forEach((item, index) => {
      item.addEventListener('pointermove', () => highlight(index));
      item.addEventListener('click', () => { active = index; commit(); trigger.focus({ preventScroll: true }); });
    });
    popup.addEventListener('toggle', () => { if (!isOpen()) reset(); });
    native.addEventListener('change', sync);
    const reposition = () => { if (isOpen()) position(); };
    window.addEventListener('resize', reposition);
    window.addEventListener('scroll', reposition, true);
    window.visualViewport?.addEventListener('resize', reposition);
    window.visualViewport?.addEventListener('scroll', reposition);
    document.body.append(popup);
    field.append(trigger);
    native.hidden = true;
    field.classList.add('select-enhanced');
    sync();
  }
}
