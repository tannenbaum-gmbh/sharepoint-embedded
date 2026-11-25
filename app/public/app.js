const containers = document.querySelector('#containers');
const status = document.querySelector('#status');

async function api(path, options) {
  const response = await fetch(path, options);
  const data = await response.json();
  if (!response.ok) throw new Error(data.error);
  return data;
}

async function showFiles() {
  const files = document.querySelector('#files');
  files.replaceChildren();
  if (!containers.value) return;
  const data = await api(`/api/files?containerId=${encodeURIComponent(containers.value)}`);
  for (const file of data.value) {
    const item = document.createElement('li');
    item.textContent = `${file.name} (${file.folder ? 'folder' : `${file.size} bytes`})`;
    files.append(item);
  }
}

async function refresh(selected = containers.value) {
  const data = await api('/api/containers');
  containers.replaceChildren();
  for (const container of data.value) {
    const option = document.createElement('option');
    option.value = container.id;
    option.textContent = container.displayName;
    containers.append(option);
  }
  if (data.value.some(container => container.id === selected)) containers.value = selected;
  await showFiles();
}

async function run(action) {
  status.textContent = 'Working…';
  const buttons = document.querySelectorAll('button');
  buttons.forEach(button => { button.disabled = true; });
  containers.disabled = true;
  try {
    await action();
    status.textContent = 'Done.';
  } catch (error) {
    status.textContent = error.message;
  } finally {
    buttons.forEach(button => { button.disabled = false; });
    containers.disabled = false;
  }
}

document.querySelector('#refresh').onclick = () => run(() => refresh());
containers.onchange = () => run(showFiles);
document.querySelector('#create').onsubmit = event => {
  event.preventDefault();
  run(async () => {
    const created = await api('/api/containers', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ displayName: document.querySelector('#name').value })
    });
    await refresh(created.id);
  });
};
document.querySelector('#upload').onsubmit = event => {
  event.preventDefault();
  run(async () => {
    const file = document.querySelector('#file').files[0];
    if (!containers.value) throw new Error('Create or select a container first.');
    if (!file || file.size > 1024 * 1024) throw new Error('Choose a file up to 1 MiB.');
    await api(`/api/files?containerId=${encodeURIComponent(containers.value)}&name=${encodeURIComponent(file.name)}`, {
      method: 'PUT',
      body: file
    });
    await showFiles();
  });
};
run(() => refresh());
