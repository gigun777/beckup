export function h(tag, attrs = {}, children = []) {
  const el = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (key.startsWith('on') && typeof value === 'function') {
      el.addEventListener(key.slice(2).toLowerCase(), value);
    } else if (value != null) {
      el.setAttribute(key, String(value));
    }
  }
  for (const child of children) {
    el.append(child instanceof Node ? child : document.createTextNode(String(child)));
  }
  return el;
}
