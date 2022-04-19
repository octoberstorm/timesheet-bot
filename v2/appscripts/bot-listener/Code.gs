/**
 * Record timesheet for each project via event in Hangouts Chat and store to Spreedsheet
 * Author: octoberstorm <storm.october@gmail.com>
 *
 * Supported actions: Checkin, Checkout, Get my today timesheet, Get all project members' timesheet
 * Commands: checkin, checkout, get, getall
 *
 */

 const SUPPORTED_CMDS = [
    'checkin',
    'checkout',
    'Checkin',
    'check-in',
    'Checkout',
  ];
  
  const CHECKIN_CMDS = ['checkin', 'check-in', 'Checkin', 'Check-in'];
  const CHECKOUT_CMDS = ['checkout', 'Checkout'];
  
  // URL to cloud function
  const checkinURL = 'https://<PROJECT_NAME>.cloudfunctions.net/TimesheetFunc';
  
  /**
   * Receive message from Hangouts
   * @param {Object} event the event object from Hangouts Chat
   */
  function onMessage(event) {
    var t = new Date(event.message.createTime.seconds * 1000);
    var text = event.message.text.replace('@Timesheet ', '');
  
    // validate command
    if (!isValidCommand(text)) {
      return {
        text:
          '<' +
          event.user.name +
          '> Invalid command, supported commands are `' +
          SUPPORTED_CMDS.join('`, `') +
          '`',
      };
    }
  
    var op = getOp(event.message.text);
  
    var name = '';
  
    if (event.space.type == 'DM') {
      name = 'You';
    } else {
      name = event.user.displayName;
    }
    var project = event.space.displayName;
  
    var formData = { project, op, name };
  
    var options = {
      method: 'post',
      payload: formData,
    };
  
    var response = UrlFetchApp.fetch(checkinURL, options);
    Logger.log('---UrlFetchApp---', response.getContentText());
  
    return {
      text:
        '<' + event.user.name + '> ' + text + ' logged for ' + name + ' at ' + t,
    };
  }
  
  /**
   * Responds to an ADDED_TO_SPACE event in Hangouts Chat.
   *
   * @param {Object} event the event object from Hangouts Chat
   */
  function onAddToSpace(event) {
    var message = '';
  
    if (event.space.type == 'DM') {
      message =
        'Thank you for adding me to a DM, ' + event.user.displayName + '!';
    } else {
      message = 'Thank you for adding me to ' + event.space.displayName;
    }
  
    if (event.message) {
      // Bot added through @mention.
      message = message + ' and you said: "' + event.message.text + '"';
    }
  
    return { text: message };
  }
  
  /**
   * Responds to a REMOVED_FROM_SPACE event in Hangouts Chat.
   *
   * @param {Object} event the event object from Hangouts Chat
   */
  function onRemoveFromSpace(event) {
    console.info('Bot removed from ', event.space.name);
  }
  
  function isValidCommand(text) {
    for (var i in SUPPORTED_CMDS) {
      if (text.includes(SUPPORTED_CMDS[i])) {
        return true;
      }
    }
  
    return false;
  }
  
  function getOp(text) {
    for (var i in CHECKIN_CMDS) {
      if (text.includes(CHECKIN_CMDS[i])) {
        return 'checkin';
      }
    }
  
    for (var i in CHECKOUT_CMDS) {
      if (text.includes(CHECKOUT_CMDS[i])) {
        return 'checkout';
      }
    }
  
    return 'unknown';
  }