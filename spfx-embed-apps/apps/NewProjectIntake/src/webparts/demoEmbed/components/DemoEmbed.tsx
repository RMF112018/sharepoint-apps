import * as React from 'react';
import styles from './DemoEmbed.module.scss';
import type { IDemoEmbedProps } from './IDemoEmbedProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { createSharePointClient, getCurrentUser, ensureListExists, createListItem, updateListItem } from '@hbi/sp-client';
import { Stack, TextField, PrimaryButton, DefaultButton, MessageBar, MessageBarType, Spinner, Icon, Persona, PersonaSize, Link, IconButton, Pivot, PivotItem } from '@fluentui/react';
import { FormRenderer } from './FormRenderer';
import { getCatalog, buildSelectExpand, spItemToForm, formToSpPayload } from '../../../data/fieldMapper';
import mappingRegistry from '../../../config/fieldMappingRegistry.json';
import syncConfig from '../../../config/sync.json';
import { enqueueOutbox, drainOutbox, type SyncResult } from '../../../data/syncEngine';
import { computeDelta, chunkDelta } from '../../../data/delta';
import { newProjectId } from '../../../data/canonical';
import fieldCatalog from '../../../config/fieldCatalog.json';
import { telemetry, nowMs, recordSyncDuration, recordSyncFailureSample } from '../../../utils/telemetry';

interface State {
  title: string;
  notes: string;
  error: string;
  success: string;
  loading: boolean;
  currentUser?: string;
  itemId?: number;
  initialForm?: Record<string, any>;
  formValues?: Record<string, any>;
  showDraftPrompt?: boolean;
  successLink?: string;
  allowProvisioning?: boolean;
  crossSiteSyncEnabled?: boolean;
  siteScopeWarning?: string;
  projectId?: string;
  registryLink?: string;
  syncStatus?: 'idle' | 'pending' | 'success' | 'error';
  syncError?: string;
  lastBackendStatus?: number;
  currentStep?: number;
  totalSteps?: number;
}

// legacy demo fields removed; mapping now handled by fieldMapper

export default class DemoEmbed extends React.Component<IDemoEmbedProps, State> {
  private static readonly NOTES_MAX = 500;
  public state: State = {
    title: '',
    notes: '',
    error: '',
    success: '',
    loading: false,
    currentUser: undefined,
    currentStep: 1,
    totalSteps: 6
  };

  private _onProjectNameChange = (projectName: string): void => {
    this.setState({ title: projectName });
  };

  private _retrySync = async (): Promise<void> => {
    this.setState({ syncStatus: 'pending', syncError: undefined });
    const res = await drainOutbox(this.props.context);
    if (!res.ok) {
      const status = res.status || 0;
      const perm = status === 401 || status === 403;
      const msg = perm ? 'No access to Preconstruction; saved locally. Your changes are queued.' : 'Network error; saved locally. Your changes are queued.';
      this.setState({ syncStatus: 'error', syncError: msg, lastBackendStatus: status });
    } else {
      this.setState({ syncStatus: 'success', lastBackendStatus: 200 });
    }
  };

  private _requestPullSync = async (): Promise<void> => {
    if (!this.state.projectId) return;
    this.setState({ syncStatus: 'pending', syncError: undefined });
    enqueueOutbox(this.props.context, {
      ProjectId: this.state.projectId,
      ChangedFields: {},
      SourceSiteUrl: this.props.context.pageContext.site.serverRelativeUrl,
      Action: 'pull'
    });
    const res = await drainOutbox(this.props.context);
    if (!res.ok) {
      const status = res.status || 0;
      const perm = status === 401 || status === 403;
      const msg = perm ? 'No access to Preconstruction; saved locally. Your request is queued.' : 'Network error; your request is queued.';
      this.setState({ syncStatus: 'error', syncError: msg, lastBackendStatus: status });
    } else {
      this.setState({ syncStatus: 'success', lastBackendStatus: 200 });
    }
  };

  private _nextStep = (): void => {
    const { currentStep, totalSteps } = this.state;
    if (currentStep! < totalSteps!) {
      this.setState({ currentStep: currentStep! + 1 });
    }
  };

  private _prevStep = (): void => {
    const { currentStep } = this.state;
    if (currentStep! > 1) {
      this.setState({ currentStep: currentStep! - 1 });
    }
  };

  private _onTabChange = (item?: PivotItem): void => {
    if (item && item.props.itemKey) {
      const stepNumber = parseInt(item.props.itemKey, 10);
      this.setState({ currentStep: stepNumber });
    }
  };

  private _canProceedToNext = (): boolean => {
    const { currentStep, title } = this.state;
    switch (currentStep) {
      case 1:
        // Check if title and required Project Info fields are filled
        if (!title.trim()) return false;
        
        // Check required fields from fieldCatalog
        const requiredProjectFields = fieldCatalog.filter((f: any) => 
          f.group === 'Project Info' && f.required
        );
        
        for (const field of requiredProjectFields) {
          const value = this.state.formValues?.[field.field_key];
          if (!value || (typeof value === 'string' && !value.trim())) {
            return false;
          }
        }
        return true;
      case 2:
        return true; // Can always proceed from step 2
      default:
        return false;
    }
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
    // Site guards
    try {
      const currentSiteRel = this.props.context.pageContext.site.serverRelativeUrl || '';
      const preconRel = (mappingRegistry as any).canonicalSiteRelativeUrl || '/sites/Preconstruction';
      const allowProvisioning = currentSiteRel.toLowerCase() === String(preconRel).toLowerCase();
      let crossSiteSyncEnabled = false;
      let siteScopeWarning: string | undefined = undefined;
      if (allowProvisioning) {
        crossSiteSyncEnabled = true; // central is allowed target/source for provisioning
      } else {
        // Treat as project site if it has the Project Facts list
        const siteFactsTitle = (mappingRegistry as any).lists?.projectSiteFactsListTitle || 'Project Facts';
        try {
          const sptest = createSharePointClient({ spfxContext: this.props.context });
          await sptest.web.lists.getByTitle(siteFactsTitle)();
          crossSiteSyncEnabled = true; // local edits may sync to Preconstruction via backend
        } catch {
          crossSiteSyncEnabled = false;
          siteScopeWarning = 'Cross-site sync disabled - local saves only';
        }
      }
      this.setState({ allowProvisioning, crossSiteSyncEnabled, siteScopeWarning });
    } catch {}
    // Edit mode if itemId present
    const url = new URL(window.location.href);
    const itemId = parseInt(url.searchParams.get('itemId') || '', 10);
    if (!isNaN(itemId)) {
      try {
        const sp = createSharePointClient({ spfxContext: this.props.context });
        const catalog = getCatalog();
        const { select, expand } = buildSelectExpand(catalog);
        const item = await sp.web.lists.getByTitle(this.props.targetListName || 'NewProjectIntake').items.getById(itemId).select(...select).expand(...expand)();
        const formState = spItemToForm(catalog as any, item);
        this.setState({ itemId, initialForm: formState, formValues: formState });
      } catch {}
    } else {
      // Restore draft
      const draft = this._readDraft();
      if (draft) this.setState({ initialForm: draft, formValues: draft, showDraftPrompt: true });
    }
    telemetry.track('load', { env: this.props.environmentMessage }, nowMs() - t0);
  }
  private _draftKey(): string {
    const user = this.props.context.pageContext.user.loginName || 'anon';
    const list = this.props.targetListName || 'NewProjectIntake';
    return `npi-draft-${list}-${user}`;
  }

  private _saveDraft = (values: Record<string, any>): void => {
    try {
      window.localStorage.setItem(this._draftKey(), JSON.stringify(values));
      this.setState({ formValues: values });
    } catch { this.setState({ formValues: values }); }
  };

  private _undoRestore = (): void => {
    this.setState({ initialForm: {}, formValues: {}, showDraftPrompt: false });
    this._saveDraft({});
  };

  private _discardDraft = (): void => {
    try { window.localStorage.removeItem(this._draftKey()); } catch {}
    this.setState({ showDraftPrompt: false });
  };

  private _copyEditLink = async (): Promise<void> => {
    if (!this.state.successLink) return;
    try { await navigator.clipboard.writeText(this.state.successLink); } catch {}
  };

  private _readDraft(): Record<string, any> | undefined {
    try {
      const raw = window.localStorage.getItem(this._draftKey());
      return raw ? JSON.parse(raw) : undefined;
    } catch { return undefined; }
  }

  // private _clearDraft(): void {
  //   try { window.localStorage.removeItem(this._draftKey()); } catch {}
  // }

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
    this.setState({ error: '', success: '', successLink: undefined });

    this.setState({ loading: true });
    const start = nowMs();
    const sp = createSharePointClient({ spfxContext: this.props.context });
    const listTitle = this.props.targetListName || 'NewProjectIntake';
    try {
      const exists = await ensureListExists(sp, listTitle);
      if (!exists) {
        this.setState({ error: '', success: '', loading: false });
        telemetry.track('submit_error', { reason: 'list_missing' }, nowMs() - start);
        return;
      }
      const catalog = getCatalog();
      const formVals = this.state.formValues || {};
      // Ensure ProjectId present
      const projectId: string = String((formVals as any).ProjectId || newProjectId());
      (formVals as any).ProjectId = projectId;
      const payload = await formToSpPayload(this.props.context, catalog as any, formVals);
      let itemId = this.state.itemId;
      if (itemId) {
        await updateListItem<any>(sp, { listTitle, id: itemId, fields: payload });
      } else {
        const res = await createListItem<any>(sp, { listTitle, fields: payload });
        itemId = res.data?.Id;
      }
      const link = `${window.location.origin}${window.location.pathname}?itemId=${itemId}`;
      // Build registry link
      const preconRel = (mappingRegistry as any).canonicalSiteRelativeUrl || '/sites/Preconstruction';
      const registryList = (mappingRegistry as any).lists?.canonicalFactsListTitle || 'Project Facts (Canonical)';
      const registryLink = `${window.location.origin}${preconRel}/Lists/${encodeURIComponent(registryList)}/AllItems.aspx?FilterField1=ProjectId&FilterValue1=${encodeURIComponent(projectId)}`;
      this.setState({ success: `Saved locally to ${listTitle}.`, loading: false, successLink: link, showDraftPrompt: false, projectId, registryLink });
      telemetry.track('localSaved', { listTitle, ProjectId: projectId });

      // Begin cross-site sync if enabled
      if (this.state.crossSiteSyncEnabled && !syncConfig.killSwitch) {
        this.setState({ syncStatus: 'pending', syncError: undefined });
        try {
          // Only sync changed fields
          const previous = this.state.initialForm || {};
          const current = this.state.formValues || {};
          // Governance: require ProjectId for cross-site operations
          const projectIdForSync: string | undefined = (current as any).ProjectId;
          if (!projectIdForSync) {
            this.setState({ syncStatus: 'error', syncError: 'Missing ProjectId. Bind this entry to the Project Registry before syncing.' });
            return;
          }
          // Enforce allowed central site for write-backs by queuing only if registry site matches
          const centralRel = (mappingRegistry as any).canonicalSiteRelativeUrl || '/sites/Preconstruction';
          if (String(centralRel).toLowerCase() !== '/sites/preconstruction') {
            // central site configurable; no-op here, backend enforces final guard
          }
          const deltaFields = computeDelta(previous, current);
          const chunks = chunkDelta(deltaFields, 100);
          for (const piece of chunks) {
            enqueueOutbox(this.props.context, {
              ProjectId: projectIdForSync,
              ChangedFields: piece,
              SourceSiteUrl: this.props.context.pageContext.site.serverRelativeUrl,
              Action: 'push'
            });
          }
          telemetry.track('syncQueued', { chunks: chunks.length, ProjectId: projectId });
          const tSyncStart = nowMs();
          telemetry.track('syncStart', { ProjectId: projectId });
          const res: SyncResult = await drainOutbox(this.props.context);
          if (!res.ok) {
            const status = res.status || 0;
            const perm = status === 401 || status === 403;
            const msg = perm ? 'No access to Preconstruction; saved locally. Your changes are queued.' : 'Network error; saved locally. Your changes are queued.';
            this.setState({ syncStatus: 'error', syncError: msg, lastBackendStatus: status });
            telemetry.track('syncError', { status, ProjectId: projectId });
            recordSyncFailureSample({ at: new Date().toISOString(), status, reason: msg });
          } else {
            this.setState({ syncStatus: 'success', lastBackendStatus: 200 });
            const dur = nowMs() - tSyncStart;
            telemetry.track('syncSuccess', { ProjectId: projectId }, dur);
            recordSyncDuration(dur);
          }
        } catch (e) {
          this.setState({ syncStatus: 'error', syncError: 'Sync failed. Will retry automatically.' });
          telemetry.track('syncError', { ProjectId: projectId });
        }
      }
      this._saveDraft({});
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
          {this.state.siteScopeWarning && (
            <MessageBar 
              messageBarType={MessageBarType.warning} 
              isMultiline={false} 
              role="status"
              className={styles.subtleNotification}
            >
              {this.state.siteScopeWarning}
            </MessageBar>
          )}
          {this.state.showDraftPrompt && (
            <MessageBar
              messageBarType={MessageBarType.info}
              isMultiline={false}
              role="status"
              className={styles.subtleNotification}
              actions={
                <Stack horizontal tokens={{ childrenGap: 4 }}>
                  <DefaultButton onClick={this._undoRestore} text="Undo" styles={{ root: { minHeight: '24px', fontSize: '11px' } }} />
                  <DefaultButton onClick={this._discardDraft} text="Discard" styles={{ root: { minHeight: '24px', fontSize: '11px' } }} />
                </Stack>
              }
            >
              Draft restored
            </MessageBar>
          )}
          {success && (
            <MessageBar messageBarType={MessageBarType.success} isMultiline={false} role="status">
              {success}
              {this.state.successLink && (
                <>
                  {' '}
                  <Link href={this.state.successLink}>Open</Link>{' '}
                  <IconButton iconProps={{ iconName: 'Copy' }} ariaLabel="Copy edit link" onClick={this._copyEditLink} />
                </>
              )}
            </MessageBar>
          )}
        </div>
        {/* Move title and subtitle above the container */}
        <div style={{ marginBottom: '16px' }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 12 }} style={{ padding: '0 16px' }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
              <Icon iconName="Edit" styles={{ root: { fontSize: 28 } }} aria-hidden="true" />
              <Stack tokens={{ childrenGap: 4 }}>
                <h1 className={styles.heroTitle}>{this.state.title ? `${this.state.title} Intake Form` : 'New Project Intake'}</h1>
                <p className={styles.heroSubtitle}>
                  {this.state.currentStep === 1 && 'Fill in the essential project details'}
                  {this.state.currentStep === 2 && 'Assign team members and manage information'}
                  {this.state.currentStep === 3 && 'Set key dates and timeline milestones'}
                  {this.state.currentStep === 4 && 'Define standard deliverables and documents'}
                  {this.state.currentStep === 5 && 'Add non-standard deliverables and custom fields'}
                  {this.state.currentStep === 6 && 'Enter project location and system access details'}
                  {!this.state.currentStep && 'Capture a project title and notes and submit to a SharePoint list.'}
                </p>
                {this.state.projectId && (
                  <span style={{ opacity: 0.8 }}>
                    ProjectId: {this.state.projectId}
                    {this.state.registryLink && (
                      <>
                        {' '}• <Link href={this.state.registryLink}>Registry</Link>
                      </>
                    )}
                  </span>
                )}
              </Stack>
            </Stack>
            {currentUser && (
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                <div className={styles.userChip} aria-label={`Current user ${currentUser}`} role="status">
                  <Persona text={currentUser} size={PersonaSize.size24} hidePersonaDetails />
                </div>
                {this.state.crossSiteSyncEnabled && (
                  <div aria-live="polite">
                    {this.state.syncStatus === 'pending' && <span title="Syncing to Preconstruction…">Syncing…</span>}
                    {this.state.syncStatus === 'success' && <span title="Last sync succeeded">In Sync</span>}
                    {this.state.syncStatus === 'error' && (
                      <>
                        <span title="Last sync failed">Sync Error</span>{' '}
                        <Link onClick={() => void this._retrySync()}>Retry</Link>
                        {' '}|{' '}
                        <Link onClick={() => void this._requestPullSync()}>Pull Canonical</Link>
                      </>
                    )}
                  </div>
                )}
              </Stack>
            )}
          </Stack>
        </div>
        
        <div className={styles.container}>
          <div className={styles.formCol}>
              <div className={styles.card}>
                {/* Step Navigation Tabs */}
                <Pivot
                  selectedKey={this.state.currentStep?.toString()}
                  onLinkClick={this._onTabChange}
                  headersOnly={true}
                  styles={{
                    root: { marginBottom: '16px' },
                    itemContainer: { padding: '0 8px' }
                  }}
                >
                  <PivotItem itemKey="1" headerText="1. Project Info" />
                  <PivotItem itemKey="2" headerText="2. Managing Info" />
                  <PivotItem itemKey="3" headerText="3. Key Dates" />
                  <PivotItem itemKey="4" headerText="4. Standard Deliverables" />
                  <PivotItem itemKey="5" headerText="5. Non-Standard Deliverables" />
                  <PivotItem itemKey="6" headerText="6. Project Details" />
                </Pivot>
                
                <div className={styles.cardHeader}>
                  Step {this.state.currentStep} of {this.state.totalSteps}
                </div>
                <form onSubmit={this.handleSubmit} aria-describedby="status-region">
                  <div id="status-region" aria-live="polite" style={{ position: 'absolute', left: -10000 }}>
                    {error ? `Error: ${error}` : success}
                  </div>
                  <Stack tokens={{ childrenGap: 16 }} className={styles.formFields}>
                    {/* Step-based form content */}
                    {this.state.currentStep === 1 && (
                      <>
                        <TextField
                          label="Project Name"
                          required
                          value={title}
                          onChange={(_, v) => this.setState({ title: v ?? '' })}
                          onBlur={() => this.setState((s) => ({ ...s }))}
                          aria-invalid={!title.trim() && !!error}
                          description="Enter the name of your project"
                          placeholder="e.g., Downtown Office Complex Renovation"
                          errorMessage={!title.trim() ? 'Project name is required.' : undefined}
                        />
                        <FormRenderer
                          context={this.props.context}
                          catalog={fieldCatalog.filter((f: any) => f.group === 'Project Info') as any}
                          initialValues={this.state.initialForm}
                          onValuesChange={this._saveDraft}
                          currentStep={1}
                          onProjectNameChange={this._onProjectNameChange}
                        />
                      </>
                    )}
                    
                    {this.state.currentStep === 2 && (
                      <>
                        <FormRenderer
                          context={this.props.context}
                          catalog={fieldCatalog.filter((f: any) => f.group === 'Managing Information') as any}
                          initialValues={this.state.initialForm}
                          onValuesChange={this._saveDraft}
                          currentStep={2}
                          onProjectNameChange={this._onProjectNameChange}
                        />
                        <TextField
                          label="Additional Notes"
                          value={notes}
                          onChange={(_, v) => this.setState({ notes: v ?? '' })}
                          onBlur={() => this.setState((s) => ({ ...s }))}
                          description={`${notes.length}/${DemoEmbed.NOTES_MAX} characters`}
                          multiline
                          rows={4}
                          placeholder="Add any additional notes or requirements for this project..."
                          errorMessage={notesTooLong ? `Please reduce to ${DemoEmbed.NOTES_MAX} characters or fewer.` : undefined}
                        />
                      </>
                    )}
                    
                    {this.state.currentStep === 3 && (
                      <FormRenderer
                        context={this.props.context}
                        catalog={fieldCatalog.filter((f: any) => f.group === 'Key Dates') as any}
                        initialValues={this.state.initialForm}
                        onValuesChange={this._saveDraft}
                        currentStep={3}
                        onProjectNameChange={this._onProjectNameChange}
                      />
                    )}
                    
                    {this.state.currentStep === 4 && (
                      <FormRenderer
                        context={this.props.context}
                        catalog={fieldCatalog.filter((f: any) => f.group === 'Final Deliverables') as any}
                        initialValues={this.state.initialForm}
                        onValuesChange={this._saveDraft}
                        currentStep={4}
                        onProjectNameChange={this._onProjectNameChange}
                      />
                    )}
                    
                    {this.state.currentStep === 5 && (
                      <FormRenderer
                        context={this.props.context}
                        catalog={fieldCatalog.filter((f: any) => f.group === 'Non-Standard Deliverables') as any}
                        initialValues={this.state.initialForm}
                        onValuesChange={this._saveDraft}
                        currentStep={5}
                        onProjectNameChange={this._onProjectNameChange}
                      />
                    )}

                    {this.state.currentStep === 6 && (
                      <FormRenderer
                        context={this.props.context}
                        catalog={fieldCatalog.filter((f: any) => f.group === 'Project Details') as any}
                        initialValues={this.state.initialForm}
                        onValuesChange={this._saveDraft}
                        currentStep={6}
                        onProjectNameChange={this._onProjectNameChange}
                      />
                    )}

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
                        {success} {this.state.successLink && (<Link href={this.state.successLink}>Copy edit link</Link>)}
                      </MessageBar>
                    )}

                    {/* Navigation buttons */}
                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                      {this.state.currentStep! > 1 && (
                        <DefaultButton 
                          text="Previous" 
                          onClick={this._prevStep}
                          iconProps={{ iconName: 'ChevronLeft' }}
                        />
                      )}
                      {this.state.currentStep! < this.state.totalSteps! && (
                        <PrimaryButton 
                          text="Next" 
                          onClick={this._nextStep}
                          disabled={!this._canProceedToNext()}
                          iconProps={{ iconName: 'ChevronRight' }}
                        />
                      )}
                      {this.state.currentStep === this.state.totalSteps && (
                        <PrimaryButton 
                          type="submit" 
                          text={loading ? 'Saving…' : (this.state.itemId ? 'Update Project' : 'Create Project')} 
                          disabled={loading || !title.trim() || notesTooLong} 
                        />
                      )}
                    </Stack>

                    {this.state.crossSiteSyncEnabled && this.state.currentStep === this.state.totalSteps && (
                      <DefaultButton
                        text={syncConfig.killSwitch ? 'Sync Disabled' : 'Resync'}
                        disabled={syncConfig.killSwitch}
                        onClick={() => void this._retrySync()}
                        title={syncConfig.killSwitch ? 'Cross-site sync is disabled by admin.' : 'Replay pending changes to Preconstruction or pull canonical snapshot (admin)'}
                      />
                    )}
                    {loading && <Spinner label="Submitting item..." />}
                  </Stack>
                </form>
              </div>
            </div>
        </div>
      </section>
    );
  }
}
