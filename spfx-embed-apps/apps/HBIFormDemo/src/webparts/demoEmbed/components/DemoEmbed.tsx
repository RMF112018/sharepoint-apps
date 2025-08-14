import * as React from 'react';
import styles from './DemoEmbed.module.scss';
import type { IDemoEmbedProps } from './IDemoEmbedProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { createSharePointClient, getCurrentUser, ensureListExists, createListItem } from '@hbi/sp-client';
import { Stack, TextField, PrimaryButton, MessageBar, MessageBarType, Spinner, Icon, Persona, PersonaSize, Link } from '@fluentui/react';
import { telemetry, nowMs } from '../../../utils/telemetry';

interface State {
  title: string;
  notes: string;
  error: string;
  success: string;
  loading: boolean;
  currentUser?: string;
}

interface DemoItemFields {
  Title: string;
  Notes: string;
}

export default class DemoEmbed extends React.Component<IDemoEmbedProps, State> {
  private static readonly NOTES_MAX = 500;
  public state: State = {
    title: '',
    notes: '',
    error: '',
    success: '',
    loading: false,
    currentUser: undefined
  };

  public async componentDidMount(): Promise<void> {
    const t0 = nowMs();
    const sp = createSharePointClient({ spfxContext: this.props.context });
    try {
      const u = await getCurrentUser(sp);
      this.setState({ currentUser: u.displayName || u.email || u.loginName });
    } catch {
      // noop
    }

    // Global key handling for a11y affordances
    window.addEventListener('keydown', this._onGlobalKeyDown, { capture: true });
    telemetry.track('load', { env: this.props.environmentMessage }, nowMs() - t0);
  }

  public componentWillUnmount(): void {
    window.removeEventListener('keydown', this._onGlobalKeyDown, true);
  }

  private _onGlobalKeyDown = (e: KeyboardEvent): void => {
    if (e.key === 'Escape' || e.key === 'Esc') {
      // Dismiss active banners/toasts
      if (this.state.error || this.state.success) {
        this.setState({ error: '', success: '' });
        e.stopPropagation();
      }
    }
    // Enter key submits via form default behavior when fields are valid
  };

  private handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    const { title, notes } = this.state;
    this.setState({ error: '', success: '' });
    if (!title.trim()) {
      this.setState({ error: 'Title is required.' });
      telemetry.track('validation_error', { field: 'title_required' });
      return;
    }
    if (notes.length > DemoEmbed.NOTES_MAX) {
      this.setState({ error: `Notes exceed the ${DemoEmbed.NOTES_MAX} character limit.` });
      telemetry.track('validation_error', { field: 'notes_max' });
      return;
    }

    this.setState({ loading: true });
    const start = nowMs();
    const sp = createSharePointClient({ spfxContext: this.props.context });
    const listTitle = this.props.targetListName || 'CursorDemo';
    try {
      const exists = await ensureListExists(sp, listTitle);
      if (!exists) {
        this.setState({ error: '', success: '', loading: false });
        telemetry.track('submit_error', { reason: 'list_missing' }, nowMs() - start);
        return;
      }
      await createListItem<DemoItemFields>(sp, { listTitle, fields: { Title: title, Notes: notes } });
      this.setState({ success: `Saved to ${listTitle}.`, title: '', notes: '', loading: false });
      setTimeout(() => this.setState({ success: '' }), 3000);
      telemetry.track('submit_success', undefined, nowMs() - start);
    } catch (err) {
      this.setState({ error: "Can't save right now. Try again or contact your admin.", loading: false });
      // eslint-disable-next-line no-console
      console.error(err);
      telemetry.track('submit_error', { reason: 'exception' }, nowMs() - start);
    }
  };
  public render(): React.ReactElement<IDemoEmbedProps> {
    const { hasTeamsContext } = this.props;
    const { title, notes, error, success, loading, currentUser } = this.state;
    const notesTooLong = notes.length > DemoEmbed.NOTES_MAX;

    return (
      <section className={`${styles.demoEmbed} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.toastContainer}>
          {success && (
            <MessageBar messageBarType={MessageBarType.success} isMultiline={false} role="status">
              {success}
            </MessageBar>
          )}
        </div>
        <div className={styles.container}>
          <div className={styles.grid}>
            <div className={styles.hero}>
              <div className={styles.welcome}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 12 }} className={styles.heroHeader}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                    <Icon iconName="Edit" styles={{ root: { fontSize: 28 } }} aria-hidden="true" />
                    <Stack tokens={{ childrenGap: 4 }}>
                      <h1 className={styles.heroTitle}>HBI Form Demo</h1>
                      <p className={styles.heroSubtitle}>Submit a title and notes to a SharePoint list.</p>
                    </Stack>
                  </Stack>
                  {currentUser && (
                    <div className={styles.userChip} aria-label={`Current user ${currentUser}`} role="status">
                      <Persona text={currentUser} size={PersonaSize.size24} hidePersonaDetails />
                    </div>
                  )}
                </Stack>
              </div>
            </div>
            <div className={styles.formCol}>
              <div className={styles.card}>
                <div className={styles.cardHeader}>Quick Entry</div>
                <form onSubmit={this.handleSubmit} aria-describedby="status-region">
                  <div id="status-region" aria-live="polite" style={{ position: 'absolute', left: -10000 }}>
                    {error ? `Error: ${error}` : success}
                  </div>
                  <Stack tokens={{ childrenGap: 16 }} className={styles.formFields}>
                    {error && (
                      <MessageBar
                        messageBarType={MessageBarType.error}
                        isMultiline
                        role="alert"
                        onDismiss={() => this.setState({ error: '' })}
                      >
                        <span>{error}</span>
                        {this.props.context.isServedFromLocalhost && (
                          <>
                            {' '}
                            <Link aria-expanded={false} aria-controls="error-details" onClick={() => this.setState({})}>View details (dev)</Link>
                          </>
                        )}
                      </MessageBar>
                    )}
                    {success && (
                      <MessageBar messageBarType={MessageBarType.success} isMultiline={false} role="status">
                        {success}
                      </MessageBar>
                    )}
                    <TextField
                      label="Title"
                      required
                      value={title}
                      onChange={(_, v) => this.setState({ title: v ?? '' })}
                      onBlur={() => this.setState((s) => ({ ...s }))}
                      aria-invalid={!title.trim() && !!error}
                      description="A short title for the item."
                      errorMessage={!title.trim() ? 'Title is required.' : undefined}
                    />
                    <TextField
                      label="Notes (optional)"
                      value={notes}
                      onChange={(_, v) => this.setState({ notes: v ?? '' })}
                      onBlur={() => this.setState((s) => ({ ...s }))}
                      description={`${notes.length}/${DemoEmbed.NOTES_MAX} characters`}
                      multiline
                      autoAdjustHeight
                      errorMessage={notesTooLong ? `Please reduce to ${DemoEmbed.NOTES_MAX} characters or fewer.` : undefined}
                    />
                    <PrimaryButton type="submit" text={loading ? 'Saving…' : 'Save entry'} disabled={loading || !title.trim() || notesTooLong} />
                    {loading && <Spinner label="Submitting item..." />}
                  </Stack>
                </form>
              </div>
            </div>
          </div>
        </div>
      </section>
    );
  }
}
