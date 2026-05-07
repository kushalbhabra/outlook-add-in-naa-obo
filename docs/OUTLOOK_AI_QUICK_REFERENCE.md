# Outlook AI Copilot - Quick Reference & Feature Checklist

## 📊 At-a-Glance Statistics

| Metric | Count |
|--------|-------|
| **Total Capabilities** | 72 |
| **Feature Categories** | 12 |
| **Priority "CRITICAL" Features** | 12 |
| **Priority "HIGH" Features** | 18 |
| **Supported Item Types** | 5 (Message, Appointment, Contact, Task, Generic) |
| **Supported Forms** | 2 (Read, Compose) |
| **Mobile-Ready Capabilities** | 25+ |
| **Graph API Integration Points** | 8 |
| **Event-Based Triggers** | 15 |

---

## 🎯 Quick Start: Top 10 Critical Features for AI Copilot

### Immediate Implementation (Week 1)

1. **✅ OnMessageSend Smart Alerts** (Event 6.1)
   - **API**: `OnMessageSend` event + `completionFunction(mode)`
   - **Use Case**: Validate email before send (compliance, tone, spam check)
   - **Manifest**: Requires `MessageSendApiEnabled` extension point
   - **Complexity**: High
   - **ROI**: HIGHEST - Blocks risky emails

2. **✅ OnNewMessageCompose** (Event 6.3)
   - **API**: `OnNewMessageCompose` event
   - **Use Case**: Auto-populate reply suggestions
   - **Complexity**: Medium
   - **ROI**: HIGH - Improves compose speed

3. **✅ Auto-Categorization** (Category 5.2)
   - **API**: `item.categories.addAsync()`
   - **Use Case**: Auto-tag emails by content
   - **Complexity**: Medium
   - **ROI**: HIGH - Saves time on manual tagging

4. **✅ Body Content Access** (Content 7.1)
   - **API**: `item.body.getAsync(CoercionType.Html)`
   - **Use Case**: Extract email content for AI analysis
   - **Complexity**: Low
   - **ROI**: CRITICAL - Foundation for all content analysis

5. **✅ Recipient Management** (Recipient 3.1)
   - **API**: `item.to.getAsync()` / `item.cc.getAsync()`
   - **Use Case**: Read & modify recipients
   - **Complexity**: Low
   - **ROI**: HIGH - Smart recipient suggestions

6. **✅ Sensitivity Labels** (Category 5.3)
   - **API**: `item.sensitivityLabel.setAsync()`
   - **Use Case**: Apply compliance labels (Confidential, Internal)
   - **Complexity**: Low
   - **ROI**: HIGH - Compliance automation

7. **✅ Attachment Scanning** (Attachment 4.6)
   - **API**: `OnAttachmentAdded` event
   - **Use Case**: Real-time security scanning
   - **Complexity**: Medium
   - **ROI**: CRITICAL - Prevents malware distribution

8. **✅ Custom Properties** (Metadata 8.1)
   - **API**: `item.customProperties.set/get()`
   - **Use Case**: Store AI analysis on item
   - **Complexity**: Low
   - **ROI**: HIGH - Persistent AI metadata

9. **✅ Microsoft Graph API Access** (Integration 11.6)
   - **API**: `getIdentityTokenAsync()` + NAA
   - **Use Case**: Query Teams, Calendar, other Graph data
   - **Complexity**: High (auth setup)
   - **ROI**: CRITICAL - Enables advanced scenarios

10. **✅ Task Pane UI** (UX 10.1)
    - **API**: `ShowTaskpane` manifest action
    - **Use Case**: Display copilot interface
    - **Complexity**: Low (UI is separate)
    - **ROI**: HIGH - User interface foundation

---

## 📋 Implementation Checklist by Phase

### Phase 1: MVP (Weeks 1-4)

- [ ] Setup manifest.json with OnMessageSend support
- [ ] Implement OnMessageSend Smart Alerts handler
- [ ] Implement OnNewMessageCompose handler
- [ ] Create backend AI validation service
- [ ] Implement basic body content reading
- [ ] Add auto-categorization logic
- [ ] Create React task pane UI
- [ ] Test with sample emails
- [ ] Document setup steps

### Phase 2: Enhanced Features (Weeks 5-8)

- [ ] Implement OnAttachmentAdded scanning
- [ ] Add attachment encryption handling
- [ ] Implement item.recipient change detection
- [ ] Add sensitivity label automation
- [ ] Setup Microsoft Graph integration (NAA auth)
- [ ] Implement meeting summary generation
- [ ] Add signature management based on account
- [ ] Create analytics dashboard
- [ ] Add email threading analysis

### Phase 3: Enterprise (Weeks 9-12)

- [ ] Multi-select bulk operations support
- [ ] Shared mailbox handling
- [ ] Delegate access support
- [ ] Mobile event-based triggers
- [ ] Compliance audit logging
- [ ] Custom regex-based activation
- [ ] Performance optimization
- [ ] Mobile UI optimization
- [ ] Enterprise deployment guide

---

## 🔑 Key Integration Points

### Event Handlers to Register

```typescript
// On Compose
Office.context.mailbox.addHandlerAsync(
  Office.EventType.MessageCompose,           // OnNewMessageCompose
  Office.EventType.MessageSend,              // OnMessageSend
  Office.EventType.AttachmentAdded,          // Real-time scanning
  Office.EventType.AttachmentRemoved,
  Office.EventType.SensitivityLabelChanged   // Label updates
);

// On Read
Office.context.mailbox.addHandlerAsync(
  Office.EventType.ItemChanged,              // Item selection changed
  Office.EventType.MessageRead                // Message opened
);

// On Appointment
Office.context.mailbox.addHandlerAsync(
  Office.EventType.AppointmentCompose,
  Office.EventType.AppointmentSend,
  Office.EventType.ItemChanged
);
```

### API Patterns for AI

| Scenario | API Pattern | Async | Example |
|----------|------------|-------|---------|
| Read current content | `item.body.getAsync()` | Yes | `Get message text for NLP` |
| Modify content | `item.body.setAsync()` | Yes | `Insert AI-generated reply` |
| Read metadata | `item.subject.getAsync()` | Yes | `Extract subject for classification` |
| Modify metadata | `item.categories.addAsync()` | Yes | `Apply auto-category` |
| Get recipients | `item.to.getAsync()` | Yes | `Validate external recipients` |
| Get attachments | `item.getAttachmentsAsync()` | Yes | `Scan for sensitive files` |
| Store AI results | `item.customProperties.set()` | No | `Store risk score` |
| Auth to Graph | `item.getIdentityTokenAsync()` | Yes | `Get token for Graph queries` |

---

## 🏗️ Manifest Essentials

### Minimum Unified Manifest

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/manifest/unified-manifest-schema.json",
  "id": "UUID-HERE",
  "version": "1.0.0.0",
  "name": "AI Copilot",
  "hosts": {
    "mail": {
      "permissions": ["mailbox.item.readWrite", "mailbox.readWrite"],
      "activationRules": [
        {
          "type": "item-compose",
          "itemType": "message",
          "runtimeId": "general",
          "enabled": true
        },
        {
          "type": "send",
          "itemType": "message",
          "runtimeId": "event-based",
          "enabled": true
        }
      ]
    }
  }
}
```

### Minimum XML Manifest

```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  <Permissions>ReadWriteMailbox</Permissions>
  <Hosts>
    <Host xsi:type="MailHost">
      <ExtensionPoint xsi:type="MessageSendApiEnabled">
        <MessageSendHandler EventLoopId="eventLoopId"/>
      </ExtensionPoint>
    </Host>
  </Hosts>
</OfficeApp>
```

---

## 🔒 Security Checklist

- [ ] Never log full email content to client console
- [ ] Always use NAA for backend authentication
- [ ] Validate token audience and issuer
- [ ] Encrypt sensitive data in custom properties
- [ ] Use HTTPS for all backend communication
- [ ] Implement rate limiting on AI service
- [ ] Audit all email modifications
- [ ] Mask PII in logs
- [ ] Implement content filtering
- [ ] Follow principle of least privilege

---

## 📊 Priority Matrix: Feature Selection

### 🔴 CRITICAL (Must Have)
- OnMessageSend Smart Alerts
- OnNewMessageCompose
- Auto-categorization
- Attachment scanning
- Body content reading
- Task pane UI
- Microsoft Graph integration
- Sensitivity labels

### 🟡 HIGH (Should Have)
- Meeting summaries
- Smart recipient suggestions
- Threading analysis
- Pinnable task pane
- Multi-select operations
- Recipient validation
- Spam reporting
- Calendar integration

### 🟢 MEDIUM (Nice to Have)
- Custom signature management
- Session data tracking
- Roaming settings
- Internet headers
- Enhanced location management
- Recurrence analysis
- Drag & drop handling

### 🔵 LOW (Future)
- Appointment logging
- Location services
- Task item support
- EWS callbacks

---

## 🔗 Dependency Tree

```
AI Copilot Foundation
├─ Task Pane UI (React/Vue)
├─ Office.js Library
├─ Event Runtime (background.js)
├─ Authentication (NAA)
│  └─ Microsoft Entra ID
│  └─ Backend Token Exchange
├─ Backend AI Service
│  ├─ NLP/Classification Models
│  ├─ LLM (GPT/Copilot)
│  └─ Database (results storage)
├─ Microsoft Graph API
│  ├─ Mail.Read
│  ├─ Calendar.Read
│  └─ User.Read
└─ Outlook API
   ├─ Compose operations
   ├─ Read operations
   └─ Event handlers
```

---

## 📱 Client Support Matrix

| Feature | Desktop | Web | Mac | Mobile | Notes |
|---------|---------|-----|-----|--------|-------|
| OnMessageSend | ✅ 1.12+ | ✅ 1.12+ | ✅ 1.12+ | ❌ | Critical for desktop |
| OnNewCompose | ✅ 1.10+ | ✅ 1.10+ | ✅ 1.10+ | ✅ 1.10+ | Mobile 4.2352+ |
| Categories | ✅ 1.8+ | ✅ 1.8+ | ✅ 1.8+ | ❌ | Compose only |
| Attachments | ✅ 1.1+ | ✅ 1.8+ | ✅ 1.1+ | ✅ 4.2352+ | Mobile limited |
| Body Access | ✅ 1.1+ | ✅ 1.1+ | ✅ 1.1+ | ✅ 1.5+ | Async required |
| Graph API | ✅ 1.12+ | ✅ 1.12+ | ✅ 1.12+ | ✅ 1.12+ | NAA required |
| Custom Props | ✅ 1.1+ | ✅ 1.1+ | ✅ 1.1+ | ❌ | Doesn't roam |

---

## ⚙️ Performance Tips

### Optimize AI Processing

- **Batch API calls**: Group multiple operations in Promise.all()
- **Cache results**: Store AI analysis in customProperties to avoid reprocessing
- **Defer heavy operations**: Use sessionData for temp processing, save only critical results
- **Implement retry logic**: Network calls can fail - always retry
- **Rate limit**: Don't overwhelm backend with simultaneous requests

### Optimize Event Handlers

```typescript
// ❌ BAD: Creates new Promise each time
Office.context.mailbox.addHandlerAsync(
  Office.EventType.ItemChanged,
  async () => {
    const content = await getContentAsync();
    const ai = await callAIAsync(content);
    // ... process
  }
);

// ✅ GOOD: Debounce rapid changes
let debounceTimeout;
Office.context.mailbox.addHandlerAsync(
  Office.EventType.ItemChanged,
  () => {
    clearTimeout(debounceTimeout);
    debounceTimeout = setTimeout(async () => {
      const content = await getContentAsync();
      const ai = await callAIAsync(content);
    }, 500); // Wait 500ms after item selection
  }
);
```

---

## 📚 Common Pitfalls & Solutions

| Pitfall | Cause | Solution |
|---------|-------|----------|
| Smart Alerts don't trigger | Manifest missing MessageSendApiEnabled | Add XML or unified manifest extension point |
| Categories not saved | Different item context | Always use getAsync in compose mode |
| Attachment access fails | Wrong permission level | Use ReadWriteItem minimum |
| Graph 401 error | NAA token invalid | Verify audience in token matches App ID URI |
| Mobile events don't fire | API not supported on mobile | Check Mailbox version for each API |
| Body.setAsync() fails | Item not in compose mode | Verify form === "compose" before calling |
| OnMessageSend blocks user | completionFunction called twice | Only call once, store result in variable |

---

## 🚀 Go-Live Checklist

### Pre-Launch (2 weeks before)

- [ ] All 4 priority use cases tested
- [ ] Edge cases handled (errors, network failures)
- [ ] Mobile testing on iOS/Android completed
- [ ] Performance benchmarks met (<2s response time)
- [ ] Security audit passed
- [ ] Compliance review passed
- [ ] Documentation complete
- [ ] Support team trained

### Launch Week

- [ ] Monitor error logs in real-time
- [ ] Track AI accuracy metrics
- [ ] Monitor backend performance
- [ ] Respond to user feedback
- [ ] Prepare rollback plan

### Post-Launch

- [ ] Weekly telemetry review
- [ ] Monthly AI model retraining
- [ ] Quarterly feature releases
- [ ] Continuous security scanning

---

## 📞 Troubleshooting Commands

### Debug OnMessageSend Issues

```typescript
// Check if event handler is registered
Object.keys(Office.context.mailbox.item).filter(k => k.includes('Handler'));

// Manually test completionFunction
const testCompletion = (mode: "block" | "send") => {
  console.log(`Would send with mode: ${mode}`);
};
testCompletion("send");

// Verify manifest registration
console.log(Office.context.mailbox.item.addHandlerAsync ? "Handler supported" : "Not supported");
```

### Debug Graph API Token Issues

```typescript
// Verify token structure
const token = await Office.context.mailbox.getIdentityTokenAsync();
const decoded = JSON.parse(atob(token.value.split('.')[1]));
console.log("Token audience:", decoded.aud);
console.log("Token issuer:", decoded.iss);

// Expected values:
// aud: "api://00e7ef2d-f553-4ffa-9076-616b18eb04c5" or GUID
// iss: "https://login.microsoftonline.com/{tenant}/v2.0"
```

---

## 📖 Reference Links

- [Official Office.js Docs](https://docs.microsoft.com/javascript/api/outlook)
- [Outlook Add-in Design Guidelines](https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-addin-design)
- [Smart Alerts Implementation](https://docs.microsoft.com/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough)
- [Microsoft Graph Mail API](https://docs.microsoft.com/graph/api/resources/message)
- [GitHub Office Add-in Samples](https://github.com/OfficeDev/Office-Add-in-samples)

---

**Quick Reference Version 1.0**  
**For use with**: OUTLOOK_AI_COPILOT_GUIDE.md & OUTLOOK_CAPABILITIES_MATRIX.csv
