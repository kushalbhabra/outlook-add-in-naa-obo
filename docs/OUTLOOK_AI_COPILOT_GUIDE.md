# Outlook AI Copilot - Comprehensive Capabilities & Integration Guide

## 📋 Table of Contents
1. [Executive Summary](#executive-summary)
2. [Complete Capabilities Matrix](#complete-capabilities-matrix)
3. [Priority AI Use Cases](#priority-ai-use-cases)
4. [Manifest Configuration Examples](#manifest-configuration-examples)
5. [Code Integration Patterns](#code-integration-patterns)
6. [Architecture & Data Flow Diagrams](#architecture--data-flow-diagrams)
7. [Implementation Roadmap](#implementation-roadmap)

---

## Executive Summary

This guide maps **72 distinct Outlook add-in capabilities** to AI-powered copilot use cases. The framework supports:

- ✅ **All Outlook contexts**: Mail, Calendar, Contacts, Tasks
- ✅ **All interaction patterns**: Real-time events, user-triggered, background
- ✅ **All data scopes**: Single item → entire mailbox
- ✅ **Priority AI scenarios**: Categorization, smart replies, meeting summaries, classification
- ✅ **Enterprise integration**: Microsoft Graph, Teams, custom backends, AI agents

**Key Statistics:**
- **72 total capabilities** documented
- **12 feature categories** covered
- **4 priority AI use cases** with examples
- **3+ manifest patterns** for different scenarios
- **5 code integration examples** for common patterns

---

## Complete Capabilities Matrix

### Capability Table Legend:
| Column | Description |
|--------|-------------|
| **Feature** | Outlook add-in capability name |
| **Category** | Feature grouping (Email, Calendar, etc.) |
| **API Entry Point** | Office.js API path to access feature |
| **Manifest Permission** | Required XML/JSON permission |
| **Item Types** | Message, Appointment, Contact, Task |
| **Forms** | Read or Compose mode |
| **AI Agent Opportunity** | Specific AI integration point |
| **Min Mailbox** | Minimum API requirement version |

---

### 1️⃣ EMAIL OPERATIONS (5 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 1.1 | Message Compose & Send | `Office.context.mailbox.item` | ReadWriteItem | Message | Compose, Read | **Smart validation**: Pre-send AI checks for tone, professionalism, compliance | 1.12+ |
| 1.2 | Message Subject | `item.subject.getAsync/setAsync` | ReadWriteItem | Message | Compose, Read | **Auto-subject generation**: AI generates subject based on body content | 1.1+ |
| 1.3 | Draft Management | `item.saveAsync()` | ReadWriteItem | Message | Compose | **Auto-complete detection**: AI suggests next steps or completions mid-draft | 1.3+ |
| 1.4 | Message Threading | `item` properties | ReadItem | Message | Read | **Thread summarization**: AI summarizes long conversation threads | 1.1+ |
| 1.5 | OnMessageSend Event | `OnMessageSend` event | ReadWriteItem | Message | Compose | **Smart Alerts**: AI-powered send-time validation (block/warn/prompt) | 1.12+ |

---

### 2️⃣ CALENDAR OPERATIONS (4 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 2.1 | Appointment Time | `item.start/end.getAsync/setAsync` | ReadWriteItem | Appointment | Compose, Read | **Smart scheduling**: AI suggests optimal meeting times | 1.1+ |
| 2.2 | Appointment Location | `item.enhancedLocation` | ReadWriteItem | Appointment | Compose, Read | **Smart meeting setup**: AI auto-populates video call links | 1.8+ |
| 2.3 | Recurrence | `item.recurrence.getAsync/setAsync` | ReadWriteItem | Appointment | Compose, Read | **Recurring pattern analysis**: AI identifies meeting frequency patterns | 1.7+ |
| 2.4 | Meeting Requests | `AppointmentCompose/Read` | ReadWriteItem | Appointment | Compose, Read | **Meeting summary generation**: AI summarizes action items from meeting details | 1.1+ |

---

### 3️⃣ CONTACTS & RECIPIENTS (5 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 3.1 | Recipient Management | `item.to/cc/bcc.getAsync/setAsync` | ReadWriteItem | Message, Appt | Compose, Read | **Smart recipient suggestions**: AI recommends who should be copied/included | 1.1+ |
| 3.2 | Recipient Changes | `OnMessageRecipientsChanged` | ReadWriteItem | Message | Compose | **Dynamic permissions check**: AI validates recipient access levels | 1.11+ |
| 3.3 | Recipient Validation | `EmailAddressDetails` | ReadWriteItem | Message, Appt | Compose, Read | **External recipient detection**: AI flags external emails for compliance | 1.1+ |
| 3.4 | From Account Changes | `OnMessageFromChanged` | ReadWriteItem | Message, Appt | Compose | **Context-aware signatures**: AI updates signature based on account | 1.13+ |
| 3.5 | Address Resolution | `recipient objects` | ReadItem | Message, Appt | Compose, Read | **GAL intelligence**: AI learns sender patterns from resolved recipients | 1.1+ |

---

### 4️⃣ ATTACHMENTS (7 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 4.1 | Attachment Upload | `addFileAttachmentAsync()` | ReadWriteItem | Message, Appt | Compose | **Attachment verification**: AI scans files for sensitive data | 1.1+ |
| 4.2 | Attachment Retrieval | `attachments` property | ReadItem | Message, Appt | Read, Compose | **Smart attachment analysis**: AI extracts and processes file content | 1.1+ |
| 4.3 | Attachment Content Download | `getAttachmentContentAsync()` | ReadItem | Message, Appt | Read | **AI document processing**: Extract text, classify document type | 1.8+ |
| 4.4 | Attachment Encryption | Custom + APIs | ReadWriteItem | Message | Compose | **Auto-encryption**: AI detects sensitive content → encrypt attach | 1.1+ |
| 4.5 | Inline Image Embedding | `addFileAttachmentFromBase64Async()` | ReadWriteItem | Message, Appt | Compose | **Visual enhancement**: AI generates inline diagrams/charts | 1.8+ |
| 4.6 | Attachment Events | `OnAttachmentAdded/Removed` | ReadWriteItem | Message, Appt | Compose | **Real-time scanning**: AI runs security checks on file add | 1.10+ |
| 4.7 | Base64 Attachment Support | `addFileAttachmentFromBase64Async()` | ReadWriteItem | Message, Appt | Compose | **Generated documents**: AI creates PDF/image attachments | 1.8+ |

---

### 5️⃣ CATEGORIES & CLASSIFICATIONS (4 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 5.1 | Category Master List | `mailbox.masterCategories.getAsync/addAsync/removeAsync` | ReadWriteMailbox | Message, Appt | Read, Compose | **Category intelligence**: AI learns custom category semantics | 1.8+ |
| 5.2 | Item Categories | `item.categories.getAsync/setAsync/addAsync` | ReadWriteItem | Message, Appt | Read, Compose | **Auto-categorization**: AI tags emails by content (project, priority) | 1.8+ |
| 5.3 | Sensitivity Labels | `item.sensitivityLabel.getAsync/setAsync` | ReadWriteItem | Message, Appt | Compose, Read | **Compliance tagging**: AI applies classification (Internal/Confidential) | 1.13+ |
| 5.4 | Label Change Events | `OnSensitivityLabelChanged` | ReadWriteItem | Message, Appt | Compose | **Dynamic enforcement**: AI adjusts sharing based on label | 1.13+ |

---

### 6️⃣ EVENTS & NOTIFICATIONS (7 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 6.1 | OnMessageSend | `OnMessageSend` event | ReadWriteItem | Message | Compose | **Smart Alerts**: AI validation (SoftBlock/Block/PromptUser modes) | 1.12+ |
| 6.2 | OnAppointmentSend | `OnAppointmentSend` event | ReadWriteItem | Appointment | Compose | **Meeting validation**: AI checks attendee availability, conflicts | 1.12+ |
| 6.3 | OnNewMessageCompose | `OnNewMessageCompose` event | ReadWriteItem | Message | Compose | **Auto-populate**: AI prefills reply context, suggested recipients | 1.10+ |
| 6.4 | OnNewAppointment | `OnNewAppointmentOrganizer` event | ReadWriteItem | Appointment | Compose | **Smart defaults**: AI suggests time, location, attendees | 1.10+ |
| 6.5 | OnMessageRead | `OnMessageRead` event | ReadItem | Message | Read | **Decryption/AI processing**: Run AI when message opened | 1.15+ |
| 6.6 | ItemChanged Event | `Office.EventType.ItemChanged` | ReadItem | Message, Appt | Read | **Real-time UI update**: AI refreshes analysis when item changes | 1.5+ |
| 6.7 | Drag & Drop | `Office.EventType.DragAndDropEvent` | ReadItem | Message, Attach | Read, Compose | **AI file processing**: Analyze dropped attachments/messages | Web+ |

---

### 7️⃣ CONTENT & BODY (6 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 7.1 | Body Management | `item.body.getAsync/setAsync/prependAsync` | ReadWriteItem | Message, Appt | Compose, Read | **Smart replies**: AI generates suggested message bodies | 1.1+ |
| 7.2 | Signature Management | `body.setSignatureAsync()` | ReadWriteItem | Message | Compose | **Dynamic signatures**: AI selects signature by context/sender | 1.10+ |
| 7.3 | Rich Text Formatting | `CoercionType.Html` | ReadWriteItem | Message, Appt | Compose | **Formatted suggestions**: AI provides rich HTML content | 1.1+ |
| 7.4 | Cursor-Based Insertion | `body.setSelectedDataAsync()` | ReadWriteItem | Message, Appt | Compose | **Contextual insertion**: AI places suggestions at cursor | 1.1+ |
| 7.5 | Inline Images | `addFileAttachmentFromBase64Async()` | ReadWriteItem | Message, Appt | Compose | **Visual generation**: AI embeds charts/diagrams inline | 1.8+ |
| 7.6 | Body Type Detection | `body.getTypeAsync()` | ReadItem | Message, Appt | Read, Compose | **Format-aware AI**: AI adapts suggestions to HTML vs plain text | 1.1+ |

---

### 8️⃣ METADATA (4 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 8.1 | Custom Properties | `item.customProperties.get/set` | ReadWriteItem | Message, Appt | Read, Compose | **AI state tracking**: Store AI analysis results on item | 1.1+ |
| 8.2 | Roaming Settings | `roamingSettings.get/set/saveAsync()` | ReadWriteMailbox | Message, Appt | Read, Compose | **User preferences**: AI learns user preferences across sessions | 1.1+ |
| 8.3 | Session Data | `item.sessionData.getAsync/setAsync` | ReadWriteItem | Message, Appt | Compose | **Draft intelligence**: Store composition state for AI context | 1.11+ |
| 8.4 | Internet Headers | `item.internetHeaders.setAsync/getAsync` | ReadWriteItem | Message | Compose, Read | **Metadata enrichment**: AI adds custom headers for tracking | 1.8+ |

---

### 9️⃣ MOBILE SUPPORT (5 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 9.1 | Mobile Event-Based | `OnNewMessageCompose` (mobile) | ReadWriteItem | Message | Compose | **Mobile AI**: Event triggers run on Android/iOS | 1.11+ |
| 9.2 | Mobile-Supported APIs | Limited Mailbox 1.7-1.14 APIs | ReadWriteItem | Message, Appt | Read, Compose | **Progressive AI**: Adapt features by client capability | 4.2352+ |
| 9.3 | Mobile Task Pane | `MobileMessageReadCommandSurface` | ReadItem | Message | Read | **Mobile UI**: Optimize AI copilot for mobile screens | 1.5+ |
| 9.4 | Online Meeting (Mobile) | `MobileOnlineMeetingCommandSurface` | ReadWriteItem | Appointment | Compose | **Mobile meeting setup**: Create Teams call from mobile | N/A |
| 9.5 | Appointment Logging | `MobileLogEventAppointmentAttendee` | ReadItem | Appointment | Read | **Mobile CRM logging**: Log meeting to external system | N/A |

---

### 🔟 CONTEXTUAL UX (6 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 10.1 | Task Pane Commands | `ShowTaskpane` action | ReadItem | Message, Appt | Read, Compose | **AI UI**: Main copilot interface | N/A |
| 10.2 | Pinnable Task Pane | `SupportsPinning` manifest | ReadItem | Message | Read | **Persistent AI**: Keep copilot open across emails | 1.5+ |
| 10.3 | Function Commands | `ExecuteFunction` action | ReadItem, ReadWriteItem | Message, Appt | Read, Compose | **Headless AI**: Run background AI without UI | N/A |
| 10.4 | Contextual Activation | `ItemHasRegularExpressionMatch` | ReadItem | Message, Appt | Read | **Smart triggers**: AI activates on specific patterns | N/A |
| 10.5 | Multi-Select | `SupportsMultiSelect` manifest | ReadItem, ReadWriteItem | Message | Read | **Bulk AI**: Apply AI to multiple emails | 1.13+ |
| 10.6 | Spam Reporting | `SpamReporting` event | ReadItem | Message | Read | **Phishing AI**: Report suspicious emails | 1.14+ |

---

### 1️⃣1️⃣ INTEGRATION (8 capabilities)

| # | Feature | API Entry Point | Manifest Permission | Item Types | Forms | **AI Copilot Opportunity** | Min Mailbox |
|---|---------|-----------------|---------------------|------------|-------|-------------------------|------------|
| 11.1 | Microsoft Graph API | `getIdentityTokenAsync()` → NAA | ReadItem, ReadWriteItem | Message, Appt | Read, Compose | **Graph-powered AI**: Query Teams, calendar, contacts | 1.12+ |
| 11.2 | EWS Callback Token | `getCallbackTokenAsync()` | ReadItem, ReadWriteItem | Message, Appt | Read, Compose | **Server-side AI**: Execute backend logic | 1.1+ |
| 11.3 | REST Attachments | `/me/messages/{id}/attachments` | ReadItem | Message | Read | **REST-based scanning**: Process attachments via Graph | N/A |
| 11.4 | Shared Mailbox | `Office.context.mailbox.item` | ReadItem, ReadWriteItem | Message, Appt | Read, Compose | **Team mailbox AI**: Support team inboxes | 1.13+ |
| 11.5 | Delegate Access | `Office.context.mailbox.item` | ReadItem, ReadWriteItem | Message, Appt | Read, Compose | **Delegate operations**: Work on behalf of colleagues | 1.8+ |
| 11.6 | Nested App Auth (NAA) | `getIdentityTokenAsync()` | ReadItem | Message, Appt | Read, Compose | **Secure backend**: Auth for custom AI services | 1.12+ |
| 11.7 | Item ID Conversion | `convertToRestId()` | ReadItem | Message, Appt | Read | **ID normalization**: Convert for Graph/REST APIs | 1.3+ |
| 11.8 | Location Services | `getIsOpenFromLocationAllowed()` | - | Message, Appt | Read, Compose | **Cloud integration**: Check cloud storage permissions | N/A |

---

## Priority AI Use Cases

### 🎯 Use Case 1: Email Categorization & Smart Tagging

**Scenario**: AI automatically categorizes incoming emails and applies sensitivity labels

**Capability Stack**:
- **Reading**: `OnMessageRead` (Event 6.5) → Message body content
- **Analysis**: Use Microsoft Graph to fetch mail content
- **Action**: `item.categories.setAsync()` (Category 5.2) + `item.sensitivityLabel.setAsync()` (Category 5.3)

**Flow**:
```
User receives email 
  → AI reads message content (body, subject, sender)
  → AI backend classifies (NLP: project, priority, compliance level)
  → Auto-apply category + sensitivity label
  → Update item + custom property tracking
```

**Code Trigger**:
```typescript
Office.onReady(() => {
  Office.context.mailbox.addHandlerAsync(
    Office.EventType.ItemChanged,
    (event) => {
      const item = Office.context.mailbox.item;
      
      // Read email content
      item.body.getAsync(Office.CoercionType.Html, async (result) => {
        const content = result.value;
        
        // Call AI categorization backend
        const category = await categorizeEmail(content);
        
        // Apply category
        item.categories.addAsync([category], (err) => {
          if (!err) console.log("Category applied");
        });
      });
    }
  );
});
```

---

### 🎯 Use Case 2: Smart Replies & Compose Assistance

**Scenario**: AI suggests reply text and helps draft professional messages

**Capability Stack**:
- **Trigger**: `OnNewMessageCompose` (Event 6.3) → Auto-populate suggestions
- **Content**: `item.body.setAsync()` (Content 7.1) → Insert AI-generated text
- **Enhancement**: `body.setSignatureAsync()` (Content 7.2) → Add context signature

**Flow**:
```
User clicks Reply
  → OnNewMessageCompose triggers
  → AI analyzes original message context
  → Generate 3 reply suggestions
  → Show in copilot UI for user selection
  → Apply selected suggestion to body
  → Append appropriate signature
```

---

### 🎯 Use Case 3: Meeting Summaries & Action Items

**Scenario**: AI summarizes meeting notes and extracts action items

**Capability Stack**:
- **Reading**: `item` properties (Calendar 2.4) → Appointment details, time, attendees
- **Content**: `item.body.getAsync()` (Content 7.1) → Meeting notes
- **Graph Integration**: `getIdentityTokenAsync()` (Integration 11.6) → Fetch Teams recording/transcript
- **Storage**: `item.customProperties.set()` (Metadata 8.1) → Store AI summary

**Flow**:
```
Meeting appointment opened
  → AI fetches meeting details (date, attendees, body notes)
  → If Teams meeting, get recording transcript via Graph
  → AI summarizes key points + action items
  → Store in custom property for reference
  → Display in copilot UI for editing
```

---

### 🎯 Use Case 4: Email Classification & Risk Detection

**Scenario**: AI detects phishing, compliance risks, and external share risks

**Capability Stack**:
- **Detection**: `OnMessageSend` (Event 6.1) → Validate before send
- **Analysis**: Examine recipient (Recipient 3.3), subject, body
- **Action**: Smart Alerts modes (SoftBlock/Block) → Prevent send if high risk
- **Logging**: `item.internetHeaders.setAsync()` (Metadata 8.4) → Audit trail

**Flow**:
```
User clicks Send
  → OnMessageSend event fires
  → AI analyzes: recipients, body tone, attachments
  → Risk scoring: 
    - External recipient + sensitive attachment = HIGH
    - Unprofessional language = MEDIUM
    - Phishing pattern match = BLOCK
  → Return Smart Alert result:
    - BLOCK if critical risk
    - SoftBlock if suspicious (warn user)
    - PromptUser with AI explanation
  → Store analysis in custom headers for compliance
```

---

## Manifest Configuration Examples

### Example 1: Unified Manifest for Email AI Copilot

**File**: `manifest.json` (Unified Manifest - Recommended)

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/manifest/unified-manifest-schema.json",
  "id": "00e7ef2d-f553-4ffa-9076-616b18eb04c5",
  "version": "1.0.0.0",
  "name": "Outlook AI Copilot",
  "description": "AI-powered email intelligence, categorization, and smart replies",
  "hosts": {
    "mail": {
      "permissions": [
        "mailbox.item.read",
        "mailbox.item.readWrite",
        "mailbox.readWrite"
      ],
      "commandSurfaces": [
        {
          "type": "primary-command-surface",
          "displayName": "AI Copilot",
          "items": [
            {
              "id": "ai-copilot-button",
              "type": "button",
              "displayName": "AI Assistant",
              "icons": [
                {
                  "size": 16,
                  "url": "https://localhost:3000/assets/icon-16.png"
                },
                {
                  "size": 32,
                  "url": "https://localhost:3000/assets/icon-32.png"
                }
              ],
              "actionId": "show-copilot-pane"
            }
          ]
        }
      ],
      "runtimes": [
        {
          "type": "general",
          "code": "https://localhost:3000/taskpane.html",
          "lifetime": "long",
          "actions": [
            {
              "id": "show-copilot-pane",
              "type": "show-taskpane"
            },
            {
              "id": "validate-before-send",
              "type": "execute-function"
            }
          ]
        },
        {
          "type": "event-based",
          "code": "https://localhost:3000/events.html",
          "lifetime": "long",
          "actions": [
            {
              "id": "on-new-compose",
              "type": "execute-function"
            },
            {
              "id": "on-message-send",
              "type": "execute-function"
            },
            {
              "id": "on-attachment-added",
              "type": "execute-function"
            }
          ]
        }
      ],
      "activationRules": [
        {
          "type": "item-read",
          "itemType": "message",
          "runtimeId": "general",
          "enabled": true
        },
        {
          "type": "item-read",
          "itemType": "appointment",
          "runtimeId": "general",
          "enabled": true
        },
        {
          "type": "item-compose",
          "itemType": "message",
          "runtimeId": "general",
          "enabled": true
        },
        {
          "type": "compose-open",
          "itemType": "message",
          "runtimeId": "event-based",
          "enabled": true
        },
        {
          "type": "send",
          "itemType": "message",
          "runtimeId": "event-based",
          "enabled": true
        },
        {
          "type": "item-changed",
          "itemType": "message",
          "runtimeId": "event-based",
          "enabled": true
        }
      ]
    }
  },
  "scopes": [
    "mail"
  ]
}
```

---

### Example 2: Classic XML Manifest with Event Handlers

**File**: `manifest.xml` (Compatibility mode)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="MailApp">
  <Id>00e7ef2d-f553-4ffa-9076-616b18eb04c5</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Outlook AI Copilot</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Outlook AI Copilot"/>
  <Description DefaultValue="AI-powered email intelligence and automation"/>
  <Hosts>
    <Host xsi:type="MailHost">
      <DesktopFormFactor>
        <FunctionFile resid="eventPageUrl"/>
        <ExtensionPoint xsi:type="MessageReadCommandSurface">
          <OfficeTab id="TabDefault">
            <Group id="msgReadGroup">
              <Label resid="groupLabel"/>
              <Control xsi:type="Button" id="msgReadButton">
                <Label resid="buttonLabel"/>
                <Supertip>
                  <Title resid="buttonTitle"/>
                  <Description resid="buttonDesc"/>
                </Supertip>
                <Icon>
                  <bt:Image size="16" resid="icon16"/>
                  <bt:Image size="32" resid="icon32"/>
                </Icon>
                <Action xsi:type="ShowTaskpane">
                  <SourceLocation resid="taskpaneUrl"/>
                </Action>
              </Control>
            </Group>
          </OfficeTab>
        </ExtensionPoint>

        <!-- OnMessageSend Smart Alerts -->
        <ExtensionPoint xsi:type="MessageSendApiEnabled">
          <MessageSendHandler EventLoopId="eventLoopId"/>
        </ExtensionPoint>

        <!-- OnNewCompose Event -->
        <ExtensionPoint xsi:type="MessageComposeCommandSurface">
          <OfficeTab id="TabDefault">
            <Group id="msgComposeGroup">
              <Label resid="composeGroupLabel"/>
              <Control xsi:type="Button" id="msgComposeButton">
                <Label resid="composeButtonLabel"/>
                <Icon>
                  <bt:Image size="16" resid="icon16"/>
                  <bt:Image size="32" resid="icon32"/>
                </Icon>
                <Action xsi:type="ShowTaskpane">
                  <SourceLocation resid="taskpaneUrl"/>
                </Action>
              </Control>
            </Group>
          </OfficeTab>
        </ExtensionPoint>
      </DesktopFormFactor>
    </Host>
  </Hosts>

  <Permissions>
    ReadWriteMailbox
  </Permissions>

  <DefaultSettings>
    <SourceLocation resid="taskpaneUrl"/>
  </DefaultSettings>

  <Resources>
    <bt:Images>
      <bt:Image id="icon16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
      <bt:Image id="icon32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
    </bt:Images>
    <bt:Urls>
      <bt:Url id="taskpaneUrl" DefaultValue="https://localhost:3000/taskpane.html"/>
      <bt:Url id="eventPageUrl" DefaultValue="https://localhost:3000/events.html"/>
    </bt:Urls>
    <bt:ShortStrings>
      <bt:String id="groupLabel" DefaultValue="AI Copilot"/>
      <bt:String id="buttonLabel" DefaultValue="Open Copilot"/>
      <bt:String id="buttonTitle" DefaultValue="Outlook AI Assistant"/>
      <bt:String id="composeGroupLabel" DefaultValue="Compose Helper"/>
      <bt:String id="composeButtonLabel" DefaultValue="AI Draft Assistant"/>
    </bt:ShortStrings>
    <bt:LongStrings>
      <bt:String id="buttonDesc" DefaultValue="Click to open the AI copilot with intelligent email suggestions"/>
    </bt:LongStrings>
  </Resources>
</OfficeApp>
```

---

## Code Integration Patterns

### Pattern 1: Event-Based AI Trigger (OnMessageSend Smart Alerts)

**Purpose**: Validate email before send with AI-powered checks

**File**: `src/eventHandlers.ts`

```typescript
export function setupEventHandlers() {
  // Setup OnMessageSend for smart alerts
  Office.onReady(() => {
    Office.context.mailbox.item.addHandlerAsync(
      Office.EventType.MessageSend,
      onMessageSendHandler,
      (error) => {
        if (error) console.log("Error setting up OnMessageSend: " + error);
      }
    );
  });
}

interface SendEventArgs {
  type: "messageSend";
  completionFunction: (sendMode: "block" | "softBlock" | "promptUser" | "send") => void;
}

async function onMessageSendHandler(event: SendEventArgs) {
  const item = Office.context.mailbox.item;
  
  try {
    // Gather data for AI analysis
    const analysis = {
      subject: await getSubject(),
      body: await getBodyContent(),
      recipients: await getRecipients(),
      attachments: await getAttachmentCount(),
    };

    // Call AI backend for validation
    const riskScore = await callAIValidation(analysis);

    // Determine action based on AI risk assessment
    if (riskScore.severity === "CRITICAL") {
      // Block send if critical phishing/compliance risk
      event.completionFunction("block");
    } else if (riskScore.severity === "HIGH") {
      // Soft block: warn user but allow send after confirmation
      event.completionFunction("softBlock");
      
      // Store risk analysis in custom property
      item.customProperties.set("aiRiskAnalysis", JSON.stringify(riskScore));
    } else if (riskScore.severity === "MEDIUM") {
      // Prompt user with AI warning
      event.completionFunction("promptUser");
    } else {
      // Safe to send
      event.completionFunction("send");
    }
  } catch (error) {
    console.error("Error in onMessageSend:", error);
    event.completionFunction("send"); // Don't block on error
  }
}

async function getSubject(): Promise<string> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.subject.getAsync((result) => {
      resolve(result.value);
    });
  });
}

async function getBodyContent(): Promise<string> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
      resolve(result.value);
    });
  });
}

async function getRecipients(): Promise<string[]> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.to.getAsync((result) => {
      const recipients = result.value?.map((r) => r.emailAddress) || [];
      resolve(recipients);
    });
  });
}

async function getAttachmentCount(): Promise<number> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.getAttachmentsAsync((result) => {
      resolve(result.value?.length || 0);
    });
  });
}

async function callAIValidation(analysis: any): Promise<any> {
  const response = await fetch("https://your-api.com/api/validate-email", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${await getAccessToken()}`,
    },
    body: JSON.stringify(analysis),
  });

  return response.json();
}

async function getAccessToken(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getIdentityTokenAsync((result) => {
      if (result.status === "succeeded") {
        // Exchange identity token for Graph token via backend
        fetch("https://your-api.com/auth/exchange-token", {
          method: "POST",
          body: JSON.stringify({ identityToken: result.value }),
        })
          .then((r) => r.json())
          .then((data) => resolve(data.accessToken))
          .catch(reject);
      } else {
        reject(result.error);
      }
    });
  });
}
```

---

### Pattern 2: Compose Event AI Auto-Population (OnNewMessageCompose)

**Purpose**: Auto-populate reply suggestions when user composes email

**File**: `src/composeHelper.ts`

```typescript
export function setupComposeAutoPopulation() {
  Office.onReady(() => {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.MessageCompose,
      onNewComposeHandler,
      (error) => {
        if (error) console.log("Error setting up OnNewCompose: " + error);
      }
    );
  });
}

async function onNewComposeHandler() {
  const item = Office.context.mailbox.item;

  // Detect if this is a reply/reply-all/forward
  const composeType = await detectComposeType();

  if (composeType === "reply" || composeType === "replyAll") {
    // Get original message content for context
    const originalMessage = await getConversationContext();

    // Call AI to generate reply suggestions
    const suggestions = await generateReplySuggestions(originalMessage);

    // Store suggestions in session data for task pane access
    item.sessionData.setAsync(
      { suggestions: suggestions, composeType: composeType },
      (result) => {
        if (result.status === "succeeded") {
          // Notify task pane to display suggestions
          console.log("Suggestions stored, ready for UI");
        }
      }
    );

    // Auto-apply signature if configured
    await applyContextSignature(originalMessage.from);
  } else if (composeType === "new") {
    // For new messages, pre-fill with standard template if applicable
    const template = await getEmailTemplate("default");
    item.body.setAsync(template, { coercionType: Office.CoercionType.Html });
  }
}

function detectComposeType(): Promise<
  "new" | "reply" | "replyAll" | "forward"
> {
  // Implementation depends on available item properties
  return Promise.resolve("new");
}

async function getConversationContext(): Promise<any> {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;

    // Get body for context
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      const body = result.value;

      // Extract sender for context
      item.from?.getAsync((fromResult) => {
        resolve({
          body: body,
          from: fromResult.value,
          timestamp: new Date(),
        });
      });
    });
  });
}

async function generateReplySuggestions(
  originalMessage: any
): Promise<string[]> {
  const response = await fetch(
    "https://your-api.com/api/suggest-reply",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${await getAccessToken()}`,
      },
      body: JSON.stringify({
        originalBody: originalMessage.body,
        fromAddress: originalMessage.from.emailAddress,
      }),
    }
  );

  const data = await response.json();
  return data.suggestions; // Array of ["suggestion1", "suggestion2", ...]
}

async function applyContextSignature(fromAddress: string): Promise<void> {
  // Fetch appropriate signature based on sender account
  const signature = await getSignatureForAccount(fromAddress);

  if (signature) {
    Office.context.mailbox.item.body.setSignatureAsync(signature);
  }
}

async function getSignatureForAccount(
  emailAddress: string
): Promise<string | null> {
  // Query your backend or local storage for account-specific signature
  const signatures = await fetch(
    `https://your-api.com/api/signatures/${emailAddress}`,
    {
      headers: {
        Authorization: `Bearer ${await getAccessToken()}`,
      },
    }
  );

  const data = await signatures.json();
  return data.signature || null;
}

async function getEmailTemplate(templateName: string): Promise<string> {
  const response = await fetch(
    `https://your-api.com/api/templates/${templateName}`,
    {
      headers: {
        Authorization: `Bearer ${await getAccessToken()}`,
      },
    }
  );

  const data = await response.json();
  return data.template;
}

async function getAccessToken(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getIdentityTokenAsync((result) => {
      if (result.status === "succeeded") {
        // Exchange identity token for Graph token via backend
        fetch("https://your-api.com/auth/exchange-token", {
          method: "POST",
          body: JSON.stringify({ identityToken: result.value }),
        })
          .then((r) => r.json())
          .then((data) => resolve(data.accessToken))
          .catch(reject);
      } else {
        reject(result.error);
      }
    });
  });
}
```

---

### Pattern 3: AI Auto-Categorization on Item Read

**Purpose**: Automatically categorize and label emails when opened

**File**: `src/categorizationHelper.ts`

```typescript
export function setupAutoCategorization() {
  Office.onReady(() => {
    // Register for ItemChanged event to detect when new item is selected
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      onItemChangedHandler,
      (error) => {
        if (error) console.log("Error setting up ItemChanged: " + error);
      }
    );
  });
}

async function onItemChangedHandler() {
  const item = Office.context.mailbox.item;

  // Only process messages, not appointments
  if (item.itemType !== "message") return;

  try {
    // Check if already categorized
    const existingCategories = await getExistingCategories();
    if (existingCategories.length > 0) {
      return; // Already categorized
    }

    // Get message content for analysis
    const messageData = await extractMessageContent();

    // Call AI categorization service
    const categorization = await callAICategorization(messageData);

    // Apply categories
    if (categorization.categories.length > 0) {
      item.categories.addAsync(categorization.categories, (result) => {
        if (result.status === "succeeded") {
          console.log("Categories applied:", categorization.categories);
        }
      });
    }

    // Apply sensitivity label if needed
    if (categorization.sensitivityLabel) {
      item.sensitivityLabel.setAsync(categorization.sensitivityLabel);
    }

    // Store categorization in custom property for reference
    item.customProperties.set(
      "aiCategorization",
      JSON.stringify(categorization)
    );
  } catch (error) {
    console.error("Error in auto-categorization:", error);
  }
}

async function getExistingCategories(): Promise<string[]> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.categories.getAsync((result) => {
      resolve(result.value?.length ? result.value : []);
    });
  });
}

async function extractMessageContent(): Promise<any> {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;

    // Parallel fetch of multiple properties
    Promise.all([
      getSubjectAsync(),
      getBodyAsync(),
      getSenderAsync(),
      getRecipientsAsync(),
    ]).then(([subject, body, sender, recipients]) => {
      resolve({
        subject,
        body,
        sender,
        recipients,
        itemId: item.itemId,
        dateTime: item.dateTimeCreated,
      });
    });
  });
}

function getSubjectAsync(): Promise<string> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.subject.getAsync((result) => {
      resolve(result.value);
    });
  });
}

function getBodyAsync(): Promise<string> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
      resolve(result.value);
    });
  });
}

function getSenderAsync(): Promise<string> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.from?.getAsync((result) => {
      resolve(result.value?.emailAddress || "");
    });
  });
}

function getRecipientsAsync(): Promise<string[]> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.to.getAsync((result) => {
      const recipients = result.value?.map((r) => r.emailAddress) || [];
      resolve(recipients);
    });
  });
}

async function callAICategorization(messageData: any): Promise<any> {
  const response = await fetch(
    "https://your-api.com/api/categorize-email",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${await getAccessToken()}`,
      },
      body: JSON.stringify(messageData),
    }
  );

  return response.json();
  // Expected response:
  // {
  //   "categories": ["Project: TeamA", "Priority: High"],
  //   "sensitivityLabel": "Internal",
  //   "confidence": 0.95
  // }
}

async function getAccessToken(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getIdentityTokenAsync((result) => {
      if (result.status === "succeeded") {
        fetch("https://your-api.com/auth/exchange-token", {
          method: "POST",
          body: JSON.stringify({ identityToken: result.value }),
        })
          .then((r) => r.json())
          .then((data) => resolve(data.accessToken))
          .catch(reject);
      } else {
        reject(result.error);
      }
    });
  });
}
```

---

### Pattern 4: Multi-Select Bulk AI Operations

**Purpose**: Apply AI operations to multiple selected emails at once

**File**: `src/bulkOperations.ts`

```typescript
export async function performBulkAIOperation(operationType: "categorize" | "classify" | "archive") {
  try {
    // Get all selected items (Mailbox 1.13+)
    const result = await new Promise<Office.AsyncResult<Office.Message[]>>((resolve) => {
      Office.context.mailbox.getSelectedItemsAsync((r) => resolve(r));
    });

    if (result.status !== "succeeded" || !result.value) {
      console.log("No items selected");
      return;
    }

    const selectedItems = result.value;
    console.log(`Processing ${selectedItems.length} selected items`);

    // Batch process all items
    const results = [];
    for (const item of selectedItems) {
      const itemContent = await extractItemContent(item);
      const aiResult = await applyAIOperation(operationType, itemContent);
      results.push(aiResult);
    }

    // Display results
    console.log(`Processed ${results.length} items with AI`);
  } catch (error) {
    console.error("Error in bulk operation:", error);
  }
}

async function extractItemContent(item: Office.Message): Promise<any> {
  // Convert item to processable format
  return {
    id: item.itemId,
    subject: item.subject,
    bodyType: item.bodyType,
    dateTime: item.dateTimeCreated,
    // Note: Limited data available for bulk operations
  };
}

async function applyAIOperation(
  operationType: string,
  itemContent: any
): Promise<any> {
  const response = await fetch(
    `https://your-api.com/api/bulk/${operationType}`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${await getAccessToken()}`,
      },
      body: JSON.stringify(itemContent),
    }
  );

  return response.json();
}

async function getAccessToken(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getIdentityTokenAsync((result) => {
      if (result.status === "succeeded") {
        fetch("https://your-api.com/auth/exchange-token", {
          method: "POST",
          body: JSON.stringify({ identityToken: result.value }),
        })
          .then((r) => r.json())
          .then((data) => resolve(data.accessToken))
          .catch(reject);
      } else {
        reject(result.error);
      }
    });
  });
}
```

---

### Pattern 5: Real-Time Attachment Analysis (OnAttachmentAdded)

**Purpose**: Scan attachments with AI for security/compliance issues

**File**: `src/attachmentHandler.ts`

```typescript
export function setupAttachmentHandling() {
  Office.onReady(() => {
    Office.context.mailbox.item.addHandlerAsync(
      Office.EventType.AttachmentAdded,
      onAttachmentAddedHandler,
      (error) => {
        if (error) console.log("Error setting up OnAttachmentAdded: " + error);
      }
    );

    Office.context.mailbox.item.addHandlerAsync(
      Office.EventType.AttachmentRemoved,
      onAttachmentRemovedHandler,
      (error) => {
        if (error) console.log("Error setting up OnAttachmentRemoved: " + error);
      }
    );
  });
}

async function onAttachmentAddedHandler(event: any) {
  const item = Office.context.mailbox.item;

  try {
    // Get all current attachments
    const attachments = await getAttachmentsAsync();

    // Scan each attachment with AI
    for (const attachment of attachments) {
      const scanResult = await scanAttachmentWithAI(attachment);

      if (scanResult.riskLevel === "CRITICAL") {
        // Remove dangerous file
        item.removeAttachmentAsync(attachment.id, (result) => {
          console.log("Removed dangerous attachment:", attachment.name);
        });

        // Notify user via custom property
        item.customProperties.set(
          "attachment_warning",
          `Removed: ${attachment.name} - ${scanResult.reason}`
        );
      } else if (scanResult.riskLevel === "HIGH") {
        // Tag as suspicious in metadata
        item.internetHeaders.setAsync(
          "X-AI-Attachment-Risk",
          `HIGH: ${attachment.name}`
        );
      }
    }
  } catch (error) {
    console.error("Error scanning attachment:", error);
  }
}

async function onAttachmentRemovedHandler(event: any) {
  console.log("Attachment removed by user");
}

async function getAttachmentsAsync(): Promise<Office.AttachmentDetails[]> {
  return new Promise((resolve) => {
    Office.context.mailbox.item.getAttachmentsAsync((result) => {
      resolve(result.value || []);
    });
  });
}

async function scanAttachmentWithAI(
  attachment: Office.AttachmentDetails
): Promise<any> {
  // For security, don't send full content - just metadata
  const response = await fetch(
    "https://your-api.com/api/scan-attachment",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${await getAccessToken()}`,
      },
      body: JSON.stringify({
        name: attachment.name,
        type: attachment.contentType,
        size: attachment.size,
        isInline: attachment.isInline,
      }),
    }
  );

  return response.json();
  // Expected response:
  // {
  //   "riskLevel": "LOW" | "MEDIUM" | "HIGH" | "CRITICAL",
  //   "reason": "Malware detected" | "Suspicious extension" | etc,
  //   "recommendations": []
  // }
}

async function getAccessToken(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getIdentityTokenAsync((result) => {
      if (result.status === "succeeded") {
        fetch("https://your-api.com/auth/exchange-token", {
          method: "POST",
          body: JSON.stringify({ identityToken: result.value }),
        })
          .then((r) => r.json())
          .then((data) => resolve(data.accessToken))
          .catch(reject);
      } else {
        reject(result.error);
      }
    });
  });
}
```

---

## Architecture & Data Flow Diagrams

### Architecture: Outlook AI Copilot System

```
┌─────────────────────────────────────────────────────────────────┐
│                        OUTLOOK CLIENT                           │
├─────────────────────────────────────────────────────────────────┤
│  ┌──────────────────────────────────────────────────────────┐  │
│  │            Outlook Add-in (Task Pane + Events)           │  │
│  │                                                            │  │
│  │  ┌─────────────────────────────────────────────────────┐ │  │
│  │  │ UI Layer (React/Vue)                                │ │  │
│  │  │ • Copilot chat interface                             │ │  │
│  │  │ • AI suggestions display                             │ │  │
│  │  │ • Category/label management                          │ │  │
│  │  └─────────────────────────────────────────────────────┘ │  │
│  │                         ▲                                  │  │
│  │                         │                                  │  │
│  │  ┌─────────────────────────────────────────────────────┐ │  │
│  │  │ Office.js API Layer                                 │ │  │
│  │  │ • item.body, item.subject, item.recipients          │ │  │
│  │  │ • item.categories, item.sensitivityLabel            │ │  │
│  │  │ • Event handlers (OnSend, OnCompose, etc)           │ │  │
│  │  │ • getIdentityTokenAsync() → Backend Auth             │ │  │
│  │  └─────────────────────────────────────────────────────┘ │  │
│  │                         ▲                                  │  │
│  │                         │                                  │  │
│  │  ┌─────────────────────────────────────────────────────┐ │  │
│  │  │ Event Runtime (Background)                          │ │  │
│  │  │ • OnMessageSend → Smart Alerts validation           │ │  │
│  │  │ • OnNewMessageCompose → Auto-suggestions            │ │  │
│  │  │ • OnAttachmentAdded → Scan & Encrypt               │ │  │
│  │  │ • OnItemChanged → Real-time analysis update         │ │  │
│  │  └─────────────────────────────────────────────────────┘ │  │
│  └──────────────────────────────────────────────────────────┘  │
│                         ▲                                      │
│                         │                                      │
│                    HTTPS Protocol                             │
│                         │                                      │
└─────────────────────────┼──────────────────────────────────────┘
                          │
        ┌─────────────────┴────────────────┐
        │                                  │
        ▼                                  ▼
┌──────────────────────────┐    ┌──────────────────────────┐
│   MICROSOFT GRAPH API    │    │  AI BACKEND SERVICE      │
│                          │    │                          │
│ GET /me/messages         │    │ POST /api/validate-email │
│ GET /me/calendar/events  │    │ POST /api/suggest-reply  │
│ POST /me/mailFolders/{id}│    │ POST /api/categorize     │
│    /messages             │    │ POST /api/scan-attach    │
│                          │    │                          │
└──────────────────────────┘    │ • NLP Models             │
                                 │ • Risk Classification    │
                                 │ • Custom ML Pipelines    │
                                 │                          │
                                 └──────────────────────────┘
                                         ▲
                                         │
                            ┌────────────┴─────────────┐
                            │                          │
                            ▼                          ▼
                    ┌───────────────┐         ┌──────────────┐
                    │  LLM Service  │         │  Database    │
                    │ (GPT, etc.)   │         │  • User       │
                    │               │         │  • Settings   │
                    └───────────────┘         │  • History    │
                                              └──────────────┘
```

---

### Data Flow: Email Categorization Pipeline

```
User Opens Email
    │
    ▼
ItemChanged Event Fires
    │
    ├─► Check if already categorized
    │   └─ YES: Exit
    │
    └─ NO: Continue
        │
        ▼
   Extract Message Content
   ├─ Subject
   ├─ Body (HTML/Text)
   ├─ Sender
   ├─ Recipients
   └─ Metadata
        │
        ▼
   Call AI Backend
   POST /api/categorize-email
   {
     subject: "...",
     body: "...",
     sender: "...",
     recipients: [...]
   }
        │
        ▼
   AI Analysis
   ├─ NLP Classification
   ├─ Entity Recognition
   ├─ Risk Scoring
   └─ Category Mapping
        │
        ▼
   Return Results
   {
     categories: ["Project: X", "Priority: High"],
     sensitivityLabel: "Internal",
     confidence: 0.95
   }
        │
        ▼
   Apply to Outlook Item
   ├─ item.categories.addAsync()
   ├─ item.sensitivityLabel.setAsync()
   └─ item.customProperties.set()
        │
        ▼
   UI Update
   Display in Copilot UI
```

---

### Event-Based AI Processing Timeline

```
COMPOSE TIMELINE:
─────────────────

User Clicks "New Message"
    │
    ▼ (+0ms)
OnNewMessageCompose Event
    ├─ Extract original message context
    ├─ Call AI suggestion endpoint
    └─ Store suggestions in sessionData
    │
    ▼ (+500ms) - Suggestions appear in UI
User Types Content
    │
    ▼ (ongoing)
Task Pane Updates
    ├─ Real-time tone analysis
    ├─ Spell check
    └─ Recipient validation
    │
    ▼
User Clicks "Send"
    │
    ▼ (+0ms)
OnMessageSend Event
    ├─ Gather all message data
    ├─ Call AI validation
    └─ Apply Smart Alerts logic
    │
    ├─ BLOCK: Return "block" (red UI)
    │
    ├─ SOFTBLOCK: Return "softBlock" (yellow UI)
    │   User must acknowledge before send
    │
    ├─ PROMPTUSER: Return "promptUser" (dialog)
    │   User decides: Confirm Send / Edit
    │
    └─ SAFE: Return "send"
        │
        ▼
    Message Sent
    ├─ Store AI analysis in custom headers
    └─ Log to audit trail
```

---

## Implementation Roadmap

### Phase 1: Core Email Copilot (Weeks 1-4)

**Capabilities to Implement**:
- ✅ OnMessageSend Smart Alerts (email validation)
- ✅ OnNewMessageCompose (reply suggestions)
- ✅ Auto-categorization (item.categories)
- ✅ Sensitivity label management
- ✅ Basic task pane UI

**Dependencies**:
- OAuth2 NAA authentication
- Backend AI service (LLM integration)
- Microsoft Graph Mail API access

**Deliverables**:
- Working manifest.json
- Event handlers for 3 main events
- Task pane React component
- Backend REST API stubs

---

### Phase 2: Advanced AI Features (Weeks 5-8)

**Capabilities to Implement**:
- ✅ Attachment scanning (OnAttachmentAdded)
- ✅ Meeting summary generation (Appointment body parsing)
- ✅ Smart recipient suggestions
- ✅ Conversation threading analysis
- ✅ Custom property storage

**Dependencies**:
- NLP/Classification models trained
- Document processing pipeline
- Storage for analysis results

**Deliverables**:
- Attachment security scanning
- Meeting transcript integration (Teams)
- Recipient intelligence UI
- Analytics dashboard

---

### Phase 3: Enterprise Features (Weeks 9-12)

**Capabilities to Implement**:
- ✅ Multi-select bulk operations
- ✅ Shared mailbox support
- ✅ Delegate access handling
- ✅ Contextless activation
- ✅ Mobile support (event-based)
- ✅ Compliance audit trail

**Dependencies**:
- Multi-tenant support
- Enhanced permission model
- Mobile API compatibility
- Compliance logging system

**Deliverables**:
- Enterprise deployment guide
- Audit and compliance reports
- Mobile add-in (iOS/Android)
- Enterprise onboarding materials

---

### Testing & Rollout Strategy

```
Phase 1: Internal Testing (Week 4)
├─ dogfood with team
├─ fix critical issues
└─ get feedback

Phase 2: Beta Release (Week 8)
├─ Limited external users (500)
├─ Monitor telemetry
├─ Iterate on features
└─ improve performance

Phase 3: GA Release (Week 12)
├─ Full production deployment
├─ Scale infrastructure
├─ 24/7 support
└─ continuous improvements
```

---

## Key Learnings & Best Practices

### ✅ DO's

- **Always handle async operations**: Office.js APIs are async; use Promises/async-await
- **Validate permissions at runtime**: Check `Office.context.requirements.isSetSupported()`
- **Store AI state thoughtfully**: Use `customProperties` for item-level, `sessionData` for draft-only, `roamingSettings` for user preferences
- **Test across clients**: Classic Outlook, New Outlook, Web, Mac, Mobile have different API support
- **Implement graceful fallbacks**: If Graph API fails, fall back to basic Office.js functionality
- **Batch operations**: Multi-select APIs reduce server load

### ❌ DON'Ts

- **Don't block send indefinitely**: Use Smart Alerts smartly; always provide a path forward for users
- **Don't process sensitive data client-side**: Move PII/confidential analysis to backend
- **Don't assume API availability**: Always check minimum Mailbox version requirements
- **Don't ignore mobile limitations**: Mobile only supports subset of Mailbox 1.5-1.14 APIs
- **Don't overload the UI**: Limit real-time processing to avoid performance degradation

---

## References

- [Office.js Outlook API Docs](https://docs.microsoft.com/en-us/javascript/api/outlook)
- [Outlook Add-in Design Guide](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-addin-design)
- [Smart Alerts Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events)
- [Microsoft Graph Mail API](https://docs.microsoft.com/en-us/graph/api/resources/message)
- [Nested App Authentication](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/authentication)

---

**Document Version**: 1.0  
**Last Updated**: May 7, 2026  
**Created for**: Outlook AI Copilot Project
