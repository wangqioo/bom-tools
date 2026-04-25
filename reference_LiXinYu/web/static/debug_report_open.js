(function () {
  const PHASES = ['pick', 'load', 'shell', 'content'];
  const timers = [];

  function reducedMotion() {
    const query = window.matchMedia?.('(prefers-reduced-motion: reduce)');
    return Boolean(query?.matches);
  }

  function clearTimers() {
    while (timers.length) {
      window.clearTimeout(timers.pop());
    }
  }

  function setStep(activeStep) {
    document.querySelectorAll('[data-report-step]').forEach((node) => {
      const step = node.dataset.reportStep;
      const activeIndex = PHASES.indexOf(activeStep);
      const nodeIndex = PHASES.indexOf(step);
      node.classList.toggle('is-active', step === activeStep);
      node.classList.toggle('is-done', nodeIndex >= 0 && nodeIndex < activeIndex);
    });
  }

  function setPhase(phase) {
    const sim = document.querySelector('[data-report-open-sim]');
    if (!sim) return;
    sim.dataset.phase = phase;
    setStep(phase);
  }

  function playOpening() {
    clearTimers();
    const sim = document.querySelector('[data-report-open-sim]');
    if (!sim) return;
    sim.classList.remove('is-replaying');
    void sim.offsetWidth;
    sim.classList.add('is-replaying');
    setPhase('pick');

    if (reducedMotion()) {
      setPhase('content');
      return;
    }

    timers.push(window.setTimeout(() => setPhase('load'), 180));
    timers.push(window.setTimeout(() => setPhase('shell'), 920));
    timers.push(window.setTimeout(() => setPhase('content'), 1420));
  }

  document.addEventListener('DOMContentLoaded', () => {
    setPhase('pick');
    document.querySelectorAll('[data-report-open-action="play"]').forEach((control) => {
      control.addEventListener('click', playOpening);
    });
    window.setTimeout(playOpening, 220);
  });
}());
