(function () {
  const form = document.getElementById('expenseForm');
  const merchantInput = document.getElementById('merchant');
  const dateInput = document.getElementById('expenseDate');
  const amountInput = document.getElementById('amount');
  const currencyInput = document.getElementById('currency');
  const categoryInput = document.getElementById('category');
  const paymentMethodInput = document.getElementById('paymentMethod');
  const notesInput = document.getElementById('notes');
  const message = document.getElementById('expenseMessage');
  const clearExpenseButton = document.getElementById('clearExpense');
  const clearListButton = document.getElementById('clearList');
  const startCameraButton = document.getElementById('startCamera');
  const capturePhotoButton = document.getElementById('capturePhoto');
  const analyzeReceiptButton = document.getElementById('analyzeReceipt');
  const receiptUpload = document.getElementById('receiptUpload');
  const cameraPreview = document.getElementById('cameraPreview');
  const receiptPreview = document.getElementById('receiptPreview');
  const emptyReceiptState = document.getElementById('emptyReceiptState');
  const receiptCanvas = document.getElementById('receiptCanvas');
  const detectedData = document.getElementById('detectedData');
  const expenseRows = document.getElementById('expenseRows');
  const expenseTotal = document.getElementById('expenseTotal');
  const expenseCount = document.getElementById('expenseCount');
  const lastCategory = document.getElementById('lastCategory');

  const expenses = [];
  let cameraStream = null;
  let hasReceiptImage = false;

  function today() {
    return new Date().toISOString().slice(0, 10);
  }

  function formatAmount(amount, currency) {
    return new Intl.NumberFormat('es-CR', {
      style: 'currency',
      currency: currency || 'CRC',
      maximumFractionDigits: currency === 'CRC' ? 0 : 2
    }).format(Number(amount || 0));
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function showMessage(text, type = 'error') {
    message.textContent = text;
    message.className = type === 'success'
      ? 'mt-4 rounded-lg border border-green-200 bg-green-50 px-4 py-3 text-sm font-semibold text-green-800'
      : 'mt-4 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm font-semibold text-red-800';
  }

  function hideMessage() {
    message.textContent = '';
    message.className = 'hidden';
  }

  function stopCamera() {
    if (!cameraStream) return;
    cameraStream.getTracks().forEach((track) => track.stop());
    cameraStream = null;
  }

  function showReceiptPreview(src) {
    receiptPreview.src = src;
    receiptPreview.classList.remove('hidden');
    cameraPreview.classList.add('hidden');
    emptyReceiptState.classList.add('hidden');
    analyzeReceiptButton.disabled = false;
    hasReceiptImage = true;
  }

  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = reject;
      image.src = src;
    });
  }

  async function buildOcrCanvas(src) {
    const image = await loadImage(src);
    const maxWidth = 1800;
    const scale = Math.min(1, maxWidth / image.naturalWidth);
    const width = Math.round(image.naturalWidth * scale);
    const height = Math.round(image.naturalHeight * scale);
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d', { willReadFrequently: true });

    canvas.width = width;
    canvas.height = height;
    context.drawImage(image, 0, 0, width, height);

    const imageData = context.getImageData(0, 0, width, height);
    const data = imageData.data;

    for (let index = 0; index < data.length; index += 4) {
      const gray = (data[index] * 0.299) + (data[index + 1] * 0.587) + (data[index + 2] * 0.114);
      const contrasted = Math.max(0, Math.min(255, ((gray - 128) * 1.55) + 128));
      data[index] = contrasted;
      data[index + 1] = contrasted;
      data[index + 2] = contrasted;
    }

    context.putImageData(imageData, 0, 0);
    return canvas;
  }

  function clearReceiptPreview() {
    stopCamera();
    receiptPreview.removeAttribute('src');
    receiptPreview.classList.add('hidden');
    cameraPreview.classList.add('hidden');
    emptyReceiptState.classList.remove('hidden');
    capturePhotoButton.disabled = true;
    analyzeReceiptButton.disabled = true;
    hasReceiptImage = false;
    detectedData.innerHTML = '<p class="rounded-lg bg-gray-50 px-4 py-3 font-semibold text-gray-400">Sin total neto ni fecha detectados.</p>';
  }

  function resetForm() {
    form.reset();
    dateInput.value = today();
    currencyInput.value = 'CRC';
    categoryInput.value = 'Comida';
    paymentMethodInput.value = 'Tarjeta';
    hideMessage();
  }

  function getExpenseFromForm() {
    return {
      id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()),
      merchant: merchantInput.value.trim(),
      date: dateInput.value || today(),
      amount: Number(amountInput.value || 0),
      currency: currencyInput.value,
      category: categoryInput.value,
      paymentMethod: paymentMethodInput.value,
      notes: notesInput.value.trim()
    };
  }

  function renderDetectedData(data) {
    const amountText = data.amount
      ? formatAmount(data.amount, data.currency)
      : 'No detectado';
    const dateText = data.date || 'No detectada';
    const rawDetails = [
      data.amountLine ? `Monto: ${data.amountLine}` : '',
      data.rawDate ? `Fecha: ${data.rawDate}` : ''
    ].filter(Boolean).join(' | ');

    detectedData.innerHTML = `
      <dl class="grid gap-3">
        <div class="grid gap-3 sm:grid-cols-2">
          <div class="rounded-lg bg-gray-50 px-4 py-3">
            <dt class="text-xs font-black uppercase tracking-widest text-gray-400">Monto total neto</dt>
            <dd class="mt-1 font-bold text-gray-800">${escapeHtml(amountText)}</dd>
          </div>
          <div class="rounded-lg bg-gray-50 px-4 py-3">
            <dt class="text-xs font-black uppercase tracking-widest text-gray-400">Fecha</dt>
            <dd class="mt-1 font-bold text-gray-800">${escapeHtml(dateText)}</dd>
          </div>
        </div>
        ${rawDetails ? `<p class="rounded-lg bg-gray-50 px-4 py-3 text-xs font-semibold text-gray-500">${escapeHtml(rawDetails)}</p>` : ''}
        <p class="text-xs font-semibold text-gray-400">Revisa el total neto y la fecha antes de agregar. Esta demo no guarda ni envia la imagen.</p>
      </dl>
    `;
  }

  async function extractReceiptData() {
    if (!window.Tesseract || !window.ReceiptParser) {
      throw new Error('El motor OCR no esta disponible. Revisa la conexion y vuelve a cargar la pagina.');
    }

    const ocrCanvas = await buildOcrCanvas(receiptPreview.src);
    const result = await window.Tesseract.recognize(ocrCanvas, 'spa+eng', {
      logger(progress) {
        if (progress.status === 'recognizing text' && Number.isFinite(progress.progress)) {
          analyzeReceiptButton.textContent = `Leyendo ${Math.round(progress.progress * 100)}%`;
        }
      }
    });
    const text = result?.data?.text || '';
    const parsed = window.ReceiptParser.parseReceiptText(text);
    const data = {
      ...parsed,
      currency: 'CRC',
      ocrText: text
    };

    if (data.amount) {
      amountInput.value = String(data.amount);
      currencyInput.value = data.currency;
    }

    if (data.date) {
      dateInput.value = data.date;
    }

    notesInput.value = 'Monto total neto y fecha detectados por OCR. Revisar antes de guardar.';
    renderDetectedData(data);
    return data;
  }

  function renderExpenses() {
    const total = expenses.reduce((sum, item) => item.currency === 'CRC' ? sum + item.amount : sum, 0);
    expenseTotal.textContent = formatAmount(total, 'CRC');
    expenseCount.textContent = String(expenses.length);
    lastCategory.textContent = expenses[0]?.category || 'Sin datos';

    if (!expenses.length) {
      expenseRows.innerHTML = '<tr><td colspan="5" class="px-5 py-4 text-sm text-gray-500">Aun no hay gastos agregados.</td></tr>';
      return;
    }

    expenseRows.innerHTML = expenses.map((expense) => `
      <tr class="hover:bg-gray-50">
        <td class="whitespace-nowrap px-5 py-4 text-sm text-gray-500">${escapeHtml(expense.date)}</td>
        <td class="whitespace-nowrap px-5 py-4 text-sm font-bold text-gray-800">${escapeHtml(expense.merchant || 'Sin comercio')}</td>
        <td class="whitespace-nowrap px-5 py-4 text-sm text-gray-600">
          <span class="inline-flex rounded-full bg-blue-50 px-2.5 py-1 text-xs font-black text-blue-700">${escapeHtml(expense.category)}</span>
        </td>
        <td class="whitespace-nowrap px-5 py-4 text-sm text-gray-600">${escapeHtml(expense.paymentMethod)}</td>
        <td class="whitespace-nowrap px-5 py-4 text-right text-sm font-black text-gray-900">${formatAmount(expense.amount, expense.currency)}</td>
      </tr>
    `).join('');
  }

  startCameraButton.addEventListener('click', async () => {
    hideMessage();

    try {
      stopCamera();
      cameraStream = await navigator.mediaDevices.getUserMedia({
        video: {
          facingMode: { ideal: 'environment' }
        },
        audio: false
      });

      cameraPreview.srcObject = cameraStream;
      cameraPreview.classList.remove('hidden');
      receiptPreview.classList.add('hidden');
      emptyReceiptState.classList.add('hidden');
      capturePhotoButton.disabled = false;
      analyzeReceiptButton.disabled = true;
      hasReceiptImage = false;
    } catch (error) {
      console.error(error);
      showMessage('No se pudo abrir la camara. Revisa permisos o usa subir factura.', 'error');
    }
  });

  capturePhotoButton.addEventListener('click', () => {
    if (!cameraStream) return;

    receiptCanvas.width = cameraPreview.videoWidth || 1280;
    receiptCanvas.height = cameraPreview.videoHeight || 720;
    receiptCanvas.getContext('2d').drawImage(cameraPreview, 0, 0, receiptCanvas.width, receiptCanvas.height);
    showReceiptPreview(receiptCanvas.toDataURL('image/jpeg', 0.92));
    stopCamera();
  });

  receiptUpload.addEventListener('change', () => {
    const file = receiptUpload.files?.[0];
    if (!file) return;

    stopCamera();
    const reader = new FileReader();
    reader.onload = () => showReceiptPreview(reader.result);
    reader.readAsDataURL(file);
  });

  analyzeReceiptButton.addEventListener('click', () => {
    if (!hasReceiptImage) return;

    analyzeReceiptButton.disabled = true;
    analyzeReceiptButton.textContent = 'Analizando...';

    extractReceiptData()
      .then((data) => {
        if (!data.amount || !data.date) {
          showMessage('No pude detectar completo el total neto y la fecha. Revisa la foto o llena los campos manualmente.', 'error');
          return;
        }

        showMessage('Monto total neto y fecha detectados. Revisa antes de agregar.', 'success');
      })
      .catch((error) => {
        console.error(error);
        showMessage(error.message || 'No se pudo leer la factura con OCR.', 'error');
      })
      .finally(() => {
      analyzeReceiptButton.disabled = false;
      analyzeReceiptButton.textContent = 'Detectar total y fecha';
      });
  });

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const expense = getExpenseFromForm();

    if (!expense.amount || expense.amount <= 0 || !expense.date) {
      showMessage('Ingresa una fecha y un monto total neto mayor a cero.', 'error');
      return;
    }

    expenses.unshift(expense);
    renderExpenses();
    showMessage('Gasto agregado a la lista temporal.', 'success');
    resetForm();
  });

  clearExpenseButton.addEventListener('click', () => {
    resetForm();
    clearReceiptPreview();
  });

  clearListButton.addEventListener('click', () => {
    expenses.splice(0, expenses.length);
    renderExpenses();
    showMessage('Lista temporal vaciada.', 'success');
  });

  window.addEventListener('beforeunload', stopCamera);
  dateInput.value = today();
  renderExpenses();
}());
