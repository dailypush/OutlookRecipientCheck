Office.onReady(() => {
  const button = document.getElementById('check-button');
  if (button) {
    button.addEventListener('click', () => {
      const status = document.getElementById('status');
      if (status) {
        status.textContent = 'Use the ribbon button in compose mode to run recipient checks.';
      }
    });
  }
});
