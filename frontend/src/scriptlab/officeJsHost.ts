import { IOfficeHost } from './interfaces';

export class OfficeJsHost {
  private static _host: IOfficeHost | null = null;

  static getHost(): IOfficeHost | null {
    if (this._host) {
      return this._host;
    }

    if (typeof Office === 'undefined' || !Office.context) {
      return null;
    }

    try {
      const host = this.detectHost();
      const platform = this.detectPlatform();
      const version = this.detectVersion();

      this._host = {
        type: host,
        platform: platform,
        version: version
      };

      return this._host;
    } catch (error) {
      console.error('Error detecting Office host:', error);
      return null;
    }
  }

  private static detectHost(): IOfficeHost['type'] {
    if (Office.context.host === Office.HostType.Excel) {
      return 'Excel';
    } else if (Office.context.host === Office.HostType.Word) {
      return 'Word';
    } else if (Office.context.host === Office.HostType.PowerPoint) {
      return 'PowerPoint';
    } else if (Office.context.host === Office.HostType.Outlook) {
      return 'Outlook';
    } else if (Office.context.host === Office.HostType.OneNote) {
      return 'OneNote';
    } else {
      return 'Excel'; // Default fallback
    }
  }

  private static detectPlatform(): IOfficeHost['platform'] {
    if (Office.context.platform === Office.PlatformType.PC) {
      return 'PC';
    } else if (Office.context.platform === Office.PlatformType.Mac) {
      return 'Mac';
    } else if (Office.context.platform === Office.PlatformType.iOS) {
      return 'iOS';
    } else if (Office.context.platform === Office.PlatformType.Android) {
      return 'Android';
    } else if (Office.context.platform === Office.PlatformType.Universal) {
      return 'Universal';
    } else if (Office.context.platform === Office.PlatformType.OfficeOnline) {
      return 'Web';
    } else {
      return 'Web'; // Default fallback
    }
  }

  private static detectVersion(): string {
    try {
      if (Office.context.diagnostics) {
        return Office.context.diagnostics.version || 'Unknown';
      }
      return 'Unknown';
    } catch (error) {
      return 'Unknown';
    }
  }

  static isOfficeReady(): boolean {
    return typeof Office !== 'undefined' && 
           Office.context !== null && 
           Office.context !== undefined;
  }

  static waitForOfficeReady(): Promise<void> {
    return new Promise((resolve) => {
      if (this.isOfficeReady()) {
        resolve();
        return;
      }

      Office.onReady(() => {
        resolve();
      });
    });
  }

  static async executeInOfficeContext<T>(
    operation: (context: Excel.RequestContext) => Promise<T>
  ): Promise<T> {
    await this.waitForOfficeReady();
    
    const host = this.getHost();
    if (!host || host.type !== 'Excel') {
      throw new Error('Excel context is required for this operation');
    }

    return Excel.run(operation);
  }
}