/*
 * Invite User from spreadsheet to Slack.
 *
 * Set up
 *   Create an application sheet in a spreadsheet.
 *     - timestamp
 *     - channel
 *     - user mail address
 *     - applicant
 *     - executed at [this script use]
 *     - result [this script use]
 *   Set Script Properties: SLACK_LEGACY_TOKEN, SPREADSHEET_ID, NOTIFICATION_ROOM
 *   Set the trigger if necessary.
 *   Write this script to GAS and import the library.
 *     See: https://github.com/soundTricker/SlackApp
 */
var slackApp = SlackApp.create(PropertiesService.getScriptProperties().getProperty('SLACK_LEGACY_TOKEN'))
var spreadsheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
var channels = []
var members = []

function main() {

  // 全チャンネルの取得
  var channelList = slackApp.channelsList()
  channels = channelList.channels

  // 全ユーザーの取得
  var userList = slackApp.usersList()
  members = userList.members

  // アナウンス用ルームの取得
  var roomName = PropertiesService.getScriptProperties().getProperty('NOTIFICATION_ROOM')
  var notificationRoomId = getChannelId(roomName)

  var sheet = spreadsheet.getActiveSheet()
  var range = sheet.getDataRange()

  for (rowNum = 1; rowNum < range.getValues().length; rowNum++) {
    var log = null
    var message = null
    var row = range.getValues()[rowNum]

    // 実行済みだったら実行しない
    var executedAt = row[4]
    if (executedAt) {
      continue
    }

    var channelName = row[1]
    var userMail = row[2]

    try {
      var channelId = getChannelId(channelName)
      var userId = getUserId(userMail)

      inviteMember(channelId, userId)

      log = 'Success'
      message = 'Success: <@' + userId + '> joined <#' + channelId + '>'
    } catch (e) {
      log = e
      message = 'Error: ' + userMail + ' try to join #' + channelName + ' <' + e + '>'
    } finally {
      sheet.getRange(rowNum + 1, 5).setValue(Utilities.formatDate(new Date, 'UTC', 'yyyy/MM/dd HH:mm:ss'))
      sheet.getRange(rowNum + 1, 6).setValue(log)
    }

    try {
      slackApp.postMessage(notificationRoomId, message, {
        username: 'Inviter',
        icon_url: 'https://liginc.co.jp/wp-content/uploads/2018/07/151728907817719100_30-150x150.png'
      })
    } catch (e) {
      // ignore
    }
  }
}

function getChannelId(channelName) {
  var channelId = null

  channels.forEach(function(channel, index) {
    if (channel.name == channelName) {
      channelId = channel.id
      return
    }
  })

  if (!channelId) {
    throw new Error("Channel is not defined.")
  }

  return channelId
}

function getUserId(userMail) {
  var userId = null

  members.forEach(function(member, index) {
    if (member.profile.email == userMail) {
      userId = member.id
      return
    }
  })

  if (!userId) {
    throw new Error("User is not exists.")
  }

  return userId
}

function inviteMember(channelId, userId) {
  var ret = slackApp.channelsInvite(channelId, userId)
  if (!ret.ok) {
    throw new Error(ret.error)
  }
}
