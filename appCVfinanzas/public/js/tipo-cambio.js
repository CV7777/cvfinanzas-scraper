(function () {
  const form = document.getElementById('exchangeCommentForm');
  const commentInput = document.getElementById('exchangeComment');
  const commentCount = document.getElementById('commentCount');
  const message = document.getElementById('exchangeCommentMessage');
  const saveButton = document.getElementById('saveExchangeComment');
  const refreshButton = document.getElementById('refreshExchangeComments');
  const rows = document.getElementById('exchangeCommentRows');

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function showMessage(text, type) {
    message.textContent = text;
    message.className = type === 'success'
      ? 'mt-4 rounded-lg border border-green-200 bg-green-50 px-4 py-3 text-sm font-semibold text-green-800'
      : 'mt-4 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm font-semibold text-red-800';
  }

  function formatDate(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value || '';

    return new Intl.DateTimeFormat('es-CR', {
      dateStyle: 'medium',
      timeStyle: 'short',
      timeZone: 'America/Costa_Rica'
    }).format(date);
  }

  function renderComments(comments) {
    if (!comments.length) {
      rows.innerHTML = '<tr><td colspan="3" class="px-5 py-8 text-center text-sm text-gray-500">Aun no hay comentarios registrados.</td></tr>';
      return;
    }

    rows.innerHTML = comments.map((item) => `
      <tr class="hover:bg-gray-50">
        <td class="min-w-80 whitespace-pre-wrap px-5 py-4 text-sm text-gray-800">${escapeHtml(item.comentario)}</td>
        <td class="whitespace-nowrap px-5 py-4 text-sm font-bold text-gray-700">${escapeHtml(item.usuario)}</td>
        <td class="whitespace-nowrap px-5 py-4 text-sm text-gray-500">${escapeHtml(formatDate(item.fecha))}</td>
      </tr>
    `).join('');
  }

  async function loadComments() {
    refreshButton.disabled = true;
    rows.innerHTML = '<tr><td colspan="3" class="px-5 py-8 text-center text-sm text-gray-500">Cargando comentarios...</td></tr>';

    try {
      const response = await fetch('/api/tipo-cambio/comentarios');
      if (response.status === 401) {
        window.location.href = '/login';
        return;
      }

      const data = await response.json();
      if (!response.ok) throw new Error(data.message || 'No se pudieron cargar los comentarios.');
      renderComments(data.comentarios || []);
    } catch (error) {
      rows.innerHTML = `<tr><td colspan="3" class="px-5 py-8 text-center text-sm text-red-700">${escapeHtml(error.message)}</td></tr>`;
    } finally {
      refreshButton.disabled = false;
    }
  }

  commentInput.addEventListener('input', () => {
    commentCount.textContent = String(commentInput.value.length);
  });

  refreshButton.addEventListener('click', loadComments);

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const comentario = commentInput.value.trim();

    if (!comentario) {
      showMessage('Ingrese un comentario antes de guardar.', 'error');
      commentInput.focus();
      return;
    }

    saveButton.disabled = true;
    saveButton.lastChild.textContent = ' Guardando...';

    try {
      const response = await fetch('/api/tipo-cambio/comentarios', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ comentario })
      });
      const data = await response.json();

      if (response.status === 401) {
        window.location.href = '/login';
        return;
      }
      if (!response.ok) throw new Error(data.message || 'No se pudo guardar el comentario.');

      form.reset();
      commentCount.textContent = '0';
      showMessage('Comentario guardado correctamente.', 'success');
      await loadComments();
      commentInput.focus();
    } catch (error) {
      showMessage(error.message, 'error');
    } finally {
      saveButton.disabled = false;
      saveButton.lastChild.textContent = ' Guardar comentario';
    }
  });

  loadComments();
}());
