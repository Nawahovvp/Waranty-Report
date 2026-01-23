
/* SaveQueue Logic - Integrated */
if (typeof SaveQueue === 'undefined') {
    window.SaveQueue = {
        queue: [],
        isProcessing: false,
        add: function(payload) {
            this.queue.push(payload);
            this.updateBadge(); // Update badge when added
            this.process();
        },
        process: async function() {
            if (this.isProcessing || this.queue.length === 0) return;
            this.isProcessing = true;
            this.updateBadge(); // Update badge when starting

            while (this.queue.length > 0) {
                const item = this.queue[0]; // Peek
                try {
                     // Simulate API call or call actual
                     if (typeof postToGAS === 'function') {
                         await postToGAS(item);
                     } else {
                         console.warn('postToGAS not found, skipping sync');
                         await new Promise(r => setTimeout(r, 1000));
                     }
                     this.queue.shift(); // Remove on success
                     this.updateBadge(); // Update badge on success (decrement)
                } catch (e) {
                    console.error('Queue sync error', e);
                    // Wait and retry or move to dead letter?
                    // For now, wait 5s and retry
                    await new Promise(r => setTimeout(r, 5000));
                }
            }
            this.isProcessing = false;
            this.updateBadge();
        },
        updateBadge: function() {
            const count = this.queue.length;
            const badge = document.getElementById('saveQueueCount');
            if (badge) {
                badge.textContent = count;
                badge.style.display = count > 0 ? 'flex' : 'none';
            }
        }
    };
}

