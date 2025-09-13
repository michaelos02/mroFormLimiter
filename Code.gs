/**
 * Phase 1: Modernized with better error handling and validation
 */

function onOpen() {
  FormApp.getUi() 
    .createMenu('Form Limiter')
    .addItem('Settings', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  try {
    const html = HtmlService.createTemplateFromFile('Page').evaluate();
    html.setTitle('Form Limiting Settings');
    html.setWidth(500);
    html.setHeight(600);
    FormApp.getUi().showModalDialog(html, 'Form Limiter Settings');
  } catch (error) {
    Logger.log('Error showing dialog: ' + error.toString());
    FormApp.getUi().alert('Error opening settings. Please try again.');
  }
}

/**
 * Get current settings from script properties
 * @returns {Object} Current settings object
 */
function getSettings() {
  console.log("got to loading")
  try {
    const props = PropertiesService.getScriptProperties().getProperties();
    Logger.log('Retrieved settings: ' + JSON.stringify(props));
    return {
      date: props.date || '',
      time: props.time || '',
      number: props.number || '',
      success: true
    };
  } catch (error) {
    Logger.log('Error getting settings: ' + error.toString());
    return {
      success: false,
      error: 'Failed to retrieve settings'
    };
  }
}

/**
 * Save form limiting settings and create appropriate triggers
 * @param {string} dateValue - Date in YYYY-MM-DD format
 * @param {string} timeValue - Time in HH:MM format  
 * @param {string} numberValue - Maximum number of responses
 * @returns {Object} Result object with success/error info
 */
function saveSettings(dateValue, timeValue, numberValue) {
  try {
    // Validate inputs
    const validation = validateInputs(dateValue, timeValue, numberValue);
    if (!validation.isValid) {
      return {
        success: false,
        error: validation.error
      };
    }

    // Clear existing triggers before setting new ones
    clearFormLimiterTriggers();
    
    // Save properties
    const properties = {};
    if (dateValue) properties.date = dateValue;
    if (timeValue) properties.time = timeValue;
    if (numberValue) properties.number = numberValue;
    
    PropertiesService.getScriptProperties().setProperties(properties);
    Logger.log('Saved properties: ' + JSON.stringify(properties));
    
    // Set up triggers based on what was provided
    let triggerMessages = [];
    
    if (dateValue && timeValue) {
      const triggerResult = setDateTimeTrigger(dateValue, timeValue);
      if (triggerResult.success) {
        triggerMessages.push('Date/time trigger set for ' + dateValue + ' at ' + timeValue);
      } else {
        return triggerResult;
      }
    }
    
    if (numberValue) {
      const triggerResult = setResponseLimitTrigger(numberValue);
      if (triggerResult.success) {
        triggerMessages.push('Response limit set to ' + numberValue);
      } else {
        return triggerResult;
      }
    }
    
    const message = triggerMessages.length > 0 
      ? 'Settings saved! ' + triggerMessages.join('. ')
      : 'Settings saved!';
      
    return {
      success: true,
      message: message
    };
    
  } catch (error) {
    Logger.log('Error saving settings: ' + error.toString());
    return {
      success: false,
      error: 'Failed to save settings: ' + error.message
    };
  }
}

/**
 * Validate user inputs
 * @param {string} dateValue 
 * @param {string} timeValue 
 * @param {string} numberValue 
 * @returns {Object} Validation result
 */
function validateInputs(dateValue, timeValue, numberValue) {
  // At least one limit must be set
  if (!dateValue && !numberValue) {
    return {
      isValid: false,
      error: 'Please set either a response limit, closing date, or both.'
    };
  }
  
  // Validate date if provided
  if (dateValue) {
    const dateObj = new Date(dateValue);
    if (isNaN(dateObj.getTime())) {
      return {
        isValid: false,
        error: 'Invalid date format.'
      };
    }
    
    // Check if date is in the past
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const inputDate = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
    
    if (inputDate < today) {
      return {
        isValid: false,
        error: 'Closing date cannot be in the past.'
      };
    }
  }
  
  // Validate time format if provided
  if (timeValue && !/^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/.test(timeValue)) {
    return {
      isValid: false,
      error: 'Invalid time format. Use HH:MM format.'
    };
  }
  
  // Validate number if provided
  if (numberValue) {
    const num = parseInt(numberValue);
    if (isNaN(num) || num < 1 || num > 10000) {
      return {
        isValid: false,
        error: 'Response limit must be a number between 1 and 10,000.'
      };
    }
  }
  
  return { isValid: true };
}

/**
 * Set up date/time based trigger
 * @param {string} dateValue 
 * @param {string} timeValue 
 * @returns {Object} Result object
 */
function setDateTimeTrigger(dateValue, timeValue) {
  try {
    const dateObj = new Date(dateValue + 'T' + timeValue);
    
    // Check if the datetime is in the past
    if (dateObj <= new Date()) {
      return {
        success: false,
        error: 'The closing date and time must be in the future.'
      };
    }
    
    ScriptApp.newTrigger('closeForm')
      .timeBased()
      .at(dateObj)
      .create();
      
    Logger.log('Created date/time trigger for: ' + dateObj);
    return { success: true };
    
  } catch (error) {
    Logger.log('Error creating date/time trigger: ' + error.toString());
    return {
      success: false,
      error: 'Failed to set closing date/time trigger.'
    };
  }
}

/**
 * Set up response count based trigger
 * @param {string} numberValue 
 * @returns {Object} Result object
 */
function setResponseLimitTrigger(numberValue) {
  try {
    ScriptApp.newTrigger('checkResponseLimit')
      .forForm(FormApp.getActiveForm())
      .onFormSubmit()
      .create();
      
    Logger.log('Created response limit trigger for: ' + numberValue + ' responses');
    return { success: true };
    
  } catch (error) {
    Logger.log('Error creating response limit trigger: ' + error.toString());
    return {
      success: false,
      error: 'Failed to set response limit trigger.'
    };
  }
}

/**
 * Check if response limit has been reached (triggered on form submit)
 * @param {Object} e - Form submit event
 */
function checkResponseLimit(e) {
  try {
    const form = e.source;
    const currentResponses = form.getResponses().length;
    const maxAllowed = parseInt(PropertiesService.getScriptProperties().getProperty('number'));
    
    Logger.log(`Response check: ${currentResponses}/${maxAllowed} responses`);
    
    if (currentResponses >= maxAllowed) {
      closeForm();
    }
    
  } catch (error) {
    Logger.log('Error checking response limit: ' + error.toString());
  }
}

/**
 * Close the form and clean up triggers
 */
function closeForm() {
  try {
    const form = FormApp.getActiveForm();
    form.setAcceptingResponses(false);
    Logger.log('Form closed successfully');
    
    // Clean up only our triggers
    clearFormLimiterTriggers();
    
  } catch (error) {
    Logger.log('Error closing form: ' + error.toString());
  }
}

/**
 * Clear only Form Limiter related triggers
 */
function clearFormLimiterTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(trigger => {
      const handlerFunction = trigger.getHandlerFunction();
      if (handlerFunction === 'closeForm' || handlerFunction === 'checkResponseLimit') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    
    Logger.log(`Cleared ${deletedCount} Form Limiter triggers`);
    
  } catch (error) {
    Logger.log('Error clearing triggers: ' + error.toString());
  }
}

/**
 * Include external files (CSS/JS) in HTML template
 * @param {string} filename 
 * @returns {string} File content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}