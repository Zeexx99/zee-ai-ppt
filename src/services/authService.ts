/* global Office */

export interface AuthResult {
  token?: string;
  error?: string;
}

export const authService = {
  /**
   * Initiates Google OAuth login via Office Dialog or popup
   */
  async signInWithGoogle(clientId: string): Promise<AuthResult> {
    if (!clientId || !clientId.trim()) {
      return { error: 'Please enter a Google Client ID in Zee AI Settings first.' };
    }

    return new Promise(resolve => {
      // Relative URL construction
      const startUrl = new URL('auth/start.html', window.location.href);
      startUrl.searchParams.set('client_id', clientId.trim());

      // If running inside Office with dialog API
      if (
        typeof Office !== 'undefined' &&
        Office.context &&
        Office.context.ui &&
        typeof Office.context.ui.displayDialogAsync === 'function'
      ) {
        Office.context.ui.displayDialogAsync(
          startUrl.toString(),
          { height: 60, width: 40, promptBeforeOpen: false },
          asyncResult => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              resolve({ error: asyncResult.error.message });
              return;
            }

            const dialog = asyncResult.value;

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
              dialog.close();
              try {
                // @ts-expect-error dialog message payload
                const data = JSON.parse(arg.message);
                if (data.type === 'ZEE_AUTH_RESULT') {
                  if (data.token) {
                    resolve({ token: data.token });
                  } else {
                    resolve({ error: data.error || 'Authentication canceled' });
                  }
                }
              } catch (e) {
                resolve({ error: 'Failed to parse authentication response' });
              }
            });

            dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
              dialog.close();
              resolve({ error: 'Sign-in window was closed.' });
            });
          }
        );
      } else {
        // Browser popup fallback
        const width = 500;
        const height = 650;
        const left = window.screenX + (window.outerWidth - width) / 2;
        const top = window.screenY + (window.outerHeight - height) / 2;

        const popup = window.open(
          startUrl.toString(),
          'ZeeAuthPopup',
          `width=${width},height=${height},left=${left},top=${top},status=0,toolbar=0`
        );

        if (!popup) {
          resolve({ error: 'Popup blocked by browser. Please allow popups for this site.' });
          return;
        }

        const handleMessage = (event: MessageEvent) => {
          if (event.data && event.data.type === 'ZEE_AUTH_RESULT') {
            window.removeEventListener('message', handleMessage);
            if (event.data.token) {
              resolve({ token: event.data.token });
            } else {
              resolve({ error: event.data.error || 'Authentication canceled' });
            }
          }
        };

        window.addEventListener('message', handleMessage);
      }
    });
  }
};
