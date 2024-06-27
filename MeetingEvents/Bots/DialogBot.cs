// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using MeetingEvents.Models;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using Polly;
using Polly.CircuitBreaker;

namespace Microsoft.BotBuilderSamples
{
    // This IBot implementation can run any type of Dialog. The use of type parameterization is to allows multiple different bots
    // to be run at different endpoints within the same project. This can be achieved by defining distinct Controller types
    // each with dependency on distinct IBot types, this way ASP Dependency Injection can glue everything together without ambiguity.
    // The ConversationState is used by the Dialog system. The UserState isn't, however, it might have been used in a Dialog implementation,
    // and the requirement is that all BotState objects are saved at the end of a turn.
    public class DialogBot<T> : TeamsActivityHandler
        where T : Dialog
    {
         static readonly Random _random = new Random();
        protected readonly ILogger _logger;
        protected readonly BotState _userState;
        protected readonly BotState _conversationState;
        protected readonly Dialog _dialog;
        private readonly string _connectionName;
        private readonly string _siteUrl;
        private readonly IStatePropertyAccessor<string> _userConfigProperty;
        private readonly string _appId;
        private readonly string _appSecret; 

        public DialogBot(ConversationState conversationState, UserState userState, T dialog, ILogger<DialogBot<T>> logger, IConfiguration configuration)
        {
            _connectionName = configuration["ConnectionName"] ?? throw new NullReferenceException("ConnectionName");
            _userState = userState ?? throw new NullReferenceException(nameof(userState));
            _conversationState = conversationState ?? throw new NullReferenceException(nameof(conversationState));
            _logger = logger;
            _dialog = dialog;
            _siteUrl = configuration["SiteUrl"] ?? throw new NullReferenceException("SiteUrl");
            _userConfigProperty = userState.CreateProperty<string>("UserConfiguration");
            _appId = configuration["MicrosoftAppId"];
            _appSecret = configuration["MicrosoftAppPassword"];
        }

        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            try
            {
                await base.OnTurnAsync(turnContext, cancellationToken);

                // After the turn is complete, persist any UserState changes.
                // Save any state changes that might have occurred during the turn.
                await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
            }
            catch (Exception ex)
            {
                Console.Write(ex);
            }
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            _logger.LogInformation("Running dialog with Message Activity.");
            await _dialog.RunAsync(turnContext, _conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken);
        }

        protected async override Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
        {
            var tokenResponse = await GetTokenResponse(turnContext, query.State, cancellationToken);
            if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.Token))
            {
                // There is no token, so the user has not signed in yet.
                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                var signInLink = await GetSignInLinkAsync(turnContext, cancellationToken).ConfigureAwait(false);

                return new MessagingExtensionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Type = "auth",
                        SuggestedActions = new MessagingExtensionSuggestedAction
                        {
                            Actions = new List<CardAction>
                                    {
                                        new CardAction
                                        {
                                            Type = ActionTypes.OpenUrl,
                                            Value = signInLink,
                                            Title = "Bot Service OAuth",
                                        },
                                    },
                        },
                    },
                };
            }

            var client = new SimpleGraphClient(tokenResponse.Token);
            var profile = await client.GetMyProfile();
            var imagelink = await client.GetPhotoAsync();
            var heroCard = new ThumbnailCard
            {
                Title = "Thumbnail Card",
                Text = $"Hello {profile.DisplayName}",
                Images = new List<CardImage> { new CardImage(imagelink) }
            };
            var attachments = new MessagingExtensionAttachment(HeroCard.ContentType, null, heroCard);
            var result = new MessagingExtensionResult("list", "result", new[] { attachments });
            return new MessagingExtensionResponse(result);
        }

        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionConfigurationQuerySettingUrlAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
        {
            // The user has requested the Messaging Extension Configuration page.  
            var escapedSettings = string.Empty;
            var userConfigSettings = await _userConfigProperty.GetAsync(turnContext, () => string.Empty);
            if (!string.IsNullOrEmpty(userConfigSettings))
            {
                escapedSettings = Uri.EscapeDataString(userConfigSettings);
            }
            return new MessagingExtensionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "config",
                    SuggestedActions = new MessagingExtensionSuggestedAction
                    {
                        Actions = new List<CardAction>
                        {
                            new CardAction
                            {
                                Type = ActionTypes.OpenUrl,
                                Value = $"{_siteUrl}/searchSettings.html?settings={escapedSettings}",
                            },
                        },
                    },
                },
            };
        }

        protected override async Task OnTeamsMessagingExtensionConfigurationSettingAsync(ITurnContext<IInvokeActivity> turnContext, JObject settings, CancellationToken cancellationToken)
        {
            // When the user submits the settings page, this event is fired.
            if (settings["state"] != null)
            {
                var userConfigSettings = settings["state"].ToString();
                await _userConfigProperty.SetAsync(turnContext, userConfigSettings, cancellationToken);
            }
        }

        private async Task<string> GetSignInLinkAsync(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
            var resource = await userTokenClient.GetSignInResourceAsync(_connectionName, turnContext.Activity as Activity, null, cancellationToken).ConfigureAwait(false);
            return resource.SignInLink;
        }

        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery action, CancellationToken cancellationToken)
        {
            var tokenResponse = await GetTokenResponse(turnContext, action.State, cancellationToken);
            if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.Token))
            {
                // There is no token, so the user has not signed in yet.
                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                var signInLink = await GetSignInLinkAsync(turnContext, cancellationToken).ConfigureAwait(false);

                return new MessagingExtensionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Type = "silentAuth",
                        SuggestedActions = new MessagingExtensionSuggestedAction
                        {
                            Actions = new List<CardAction>
                                {
                                    new CardAction
                                    {
                                        Type = ActionTypes.OpenUrl,
                                        Value = signInLink,
                                        Title = "Bot Service OAuth",
                                    },
                                },
                        },
                    },
                };
            }
            var client = new SimpleGraphClient(tokenResponse.Token);
            var me = await client.GetMyProfile();
            var imagelink = await client.GetPhotoAsync();
            var previewcard = new ThumbnailCard
            {
                Title = me.DisplayName,
                Images = new List<CardImage> { new CardImage { Url = imagelink } }
            };
            var attachment = new MessagingExtensionAttachment
            {
                ContentType = ThumbnailCard.ContentType,
                Content = previewcard,
                Preview = previewcard.ToAttachment()
            };
            return new MessagingExtensionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "result",
                    AttachmentLayout = "list",
                    Attachments = new List<MessagingExtensionAttachment> { attachment }
                }
            };
        }


        protected override Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            // This method is to handle the 'Close' button on the confirmation Task Module after the user signs out.
            return Task.FromResult(new MessagingExtensionActionResponse());
        }

        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            if (action.CommandId.ToUpper() == "SHOWPROFILE")
            {
                var state = action.State; // Check the state value
                var tokenResponse = await GetTokenResponse(turnContext, state, cancellationToken);
                if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.Token))
                {
                    // There is no token, so the user has not signed in yet.

                    // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                    var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
                    var resource = await userTokenClient.GetSignInResourceAsync(_connectionName, turnContext.Activity as Activity, null, cancellationToken);

                    return new MessagingExtensionActionResponse
                    {
                        ComposeExtension = new MessagingExtensionResult
                        {
                            Type = "silentAuth",
                            SuggestedActions = new MessagingExtensionSuggestedAction
                            {
                                Actions = new List<CardAction>
                                {
                                    new CardAction
                                    {
                                        Type = ActionTypes.OpenUrl,
                                        Value = resource.SignInLink,
                                        Title = "Bot Service OAuth",
                                    },
                                },
                            },
                        },
                    };
                }
                var client = new SimpleGraphClient(tokenResponse.Token);
                var profile = await client.GetMyProfile();
                var imagelink = _siteUrl +  await client.GetPublicURLForProfilePhoto(profile.Id);
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Card = GetProfileCard(profile, imagelink),
                            Height = 250,
                            Width = 400,
                            Title = "Adaptive Card: Inputs",
                        },
                    },
                };
            }
            if (action.CommandId.ToUpper() == "SIGNOUTCOMMAND")
            {
                var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
                await userTokenClient.SignOutUserAsync(turnContext.Activity.From.Id, _connectionName, turnContext.Activity.ChannelId, cancellationToken);

                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Card = new Microsoft.Bot.Schema.Attachment
                            {
                                Content = new AdaptiveCard(new AdaptiveSchemaVersion("1.0"))
                                {
                                    Body = new List<AdaptiveElement>() { new AdaptiveTextBlock() { Text = "You have been signed out." } },
                                    Actions = new List<AdaptiveAction>() { new AdaptiveSubmitAction() { Title = "Close" } },
                                },
                                ContentType = AdaptiveCard.ContentType,
                            },
                            Height = 200,
                            Width = 400,
                            Title = "Adaptive Card: Inputs",
                        },
                    },
                };
            }
            return null;
        }

        private async Task<TokenResponse> GetTokenResponse(ITurnContext<IInvokeActivity> turnContext, string state, CancellationToken cancellationToken)
        {
            var magicCode = string.Empty;

            if (!string.IsNullOrEmpty(state))
            {
                if (int.TryParse(state, out var parsed))
                {
                    magicCode = parsed.ToString();
                }
            }

            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
            var tokenResponse = await userTokenClient.GetUserTokenAsync(turnContext.Activity.From.Id, _connectionName, turnContext.Activity.ChannelId, magicCode, cancellationToken).ConfigureAwait(false);

            return tokenResponse;
        }

        private async Task<TokenResponse> GetTokenResponse(ITurnContext<IEventActivity> turnContext, string state, CancellationToken cancellationToken)
        {
            var magicCode = string.Empty;

            if (!string.IsNullOrEmpty(state))
            {
                if (int.TryParse(state, out var parsed))
                {
                    magicCode = parsed.ToString();
                }
            }

            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
            var tokenResponse = await userTokenClient.GetUserTokenAsync(turnContext.Activity.From.Id, _connectionName, turnContext.Activity.ChannelId, magicCode, cancellationToken).ConfigureAwait(false);

            return tokenResponse;
        }

        protected override async Task<InvokeResponse> OnInvokeActivityAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            JObject valueObject = JObject.FromObject(turnContext.Activity.Value);
            if (valueObject["authentication"] != null)
            {
                JObject authenticationObject = JObject.FromObject(valueObject["authentication"]);
                if (authenticationObject["token"] != null)
                {
                    //If the token is NOT exchangeable, then return 412 to require user consent
                    if (await TokenIsExchangeable(turnContext, cancellationToken))
                    {
                        return await base.OnInvokeActivityAsync(turnContext, cancellationToken).ConfigureAwait(false);
                    }
                    else
                    {
                        var response = new InvokeResponse();
                        response.Status = 412;
                        return response;
                    }
                }
            }
            return await base.OnInvokeActivityAsync(turnContext, cancellationToken).ConfigureAwait(false);
        }

        private async Task<bool> TokenIsExchangeable(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            TokenResponse tokenExchangeResponse = null;
            try
            {
                JObject valueObject = JObject.FromObject(turnContext.Activity.Value);
                var tokenExchangeRequest =
                ((JObject)valueObject["authentication"])?.ToObject<TokenExchangeInvokeRequest>();

                var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();

                tokenExchangeResponse = await userTokenClient.ExchangeTokenAsync(turnContext.Activity.From.Id,
                    _connectionName, turnContext.Activity.ChannelId,
                    new TokenExchangeRequest { Token = tokenExchangeRequest.Token },
                    cancellationToken);
            }
#pragma warning disable CA1031 //Do not catch general exception types (ignoring, see comment below)
            catch
#pragma warning restore CA1031 //Do not catch general exception types
            {
                //ignore exceptions
                //if token exchange failed for any reason, tokenExchangeResponse above remains null, and a failure invoke response is sent to the caller.
                //This ensures the caller knows that the invoke has failed.
            }
            if (tokenExchangeResponse == null || string.IsNullOrEmpty(tokenExchangeResponse.Token))
            {
                return false;
            }
            return true;
        }

        private static Microsoft.Bot.Schema.Attachment GetProfileCard(Graph.User profile, string imagelink)
        {
            var card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0));

            card.Body.Add(new AdaptiveTextBlock()
            {
                Text = $"Hello, {profile.DisplayName}",
                Size = AdaptiveTextSize.ExtraLarge
            });

            card.Body.Add(new AdaptiveImage()
            {
                Url = new Uri(imagelink)
            });
            return new Microsoft.Bot.Schema.Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }


        /// <summary>
        /// Activity Handler for Meeting Participant join event
        /// </summary>
        /// <param name="meeting"></param>
        /// <param name="turnContext"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        protected override async Task OnTeamsMeetingParticipantsJoinAsync(MeetingParticipantsEventDetails meeting, ITurnContext<IEventActivity> turnContext, CancellationToken cancellationToken)
        {
            await turnContext.SendActivityAsync(MessageFactory.Attachment(createAdaptiveCardInvokeResponseAsync(meeting.Members[0].User.Name, " has joined the meeting.")));
            return;
        }

        /// <summary>
        /// Activity Handler for Meeting Participant leave event
        /// </summary>
        /// <param name="meeting"></param>
        /// <param name="turnContext"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        protected override async Task OnTeamsMeetingParticipantsLeaveAsync(MeetingParticipantsEventDetails meeting, ITurnContext<IEventActivity> turnContext, CancellationToken cancellationToken)
        {
            await turnContext.SendActivityAsync(MessageFactory.Attachment(createAdaptiveCardInvokeResponseAsync(meeting.Members[0].User.Name, " left the meeting.")));
            return;
        }

        /// <summary>
        /// Sample Adaptive card for Meeting participant events.
        /// </summary>
        private Bot.Schema.Attachment createAdaptiveCardInvokeResponseAsync(string userName, string action)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion("1.4"))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveRichTextBlock
                    {
                        Inlines = new List<AdaptiveInline>
                        {
                            new AdaptiveTextRun
                            {
                                Text = userName,
                                Weight = AdaptiveTextWeight.Bolder,
                                Size = AdaptiveTextSize.Default,
                            },
                            new AdaptiveTextRun
                            {
                                Text = action,
                                Weight = AdaptiveTextWeight.Default,
                                Size = AdaptiveTextSize.Default,
                            }
                        },
                    Spacing = AdaptiveSpacing.Medium,
                    }
                }
            };

            return new Bot.Schema.Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }

        /// <summary>
        /// Activity Handler for Meeting start event
        /// </summary>
        /// <param name="meeting"></param>
        /// <param name="turnContext"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        protected override async Task OnTeamsMeetingStartAsync(MeetingStartEventDetails meeting, ITurnContext<IEventActivity> turnContext, CancellationToken cancellationToken)
        {
            // Save any state changes that might have occurred during the turn.
            var conversationStateAccessors = _conversationState.CreateProperty<MeetingData>(nameof(MeetingData));
            var conversationData = await conversationStateAccessors.GetAsync(turnContext, () => new MeetingData());
            conversationData.StartTime = meeting.StartTime;
            await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
            await turnContext.SendActivityAsync(MessageFactory.Attachment(GetAdaptiveCardForMeetingStart(meeting, conversationData, turnContext, cancellationToken)));
        }

        /// <summary>
        /// Activity Handler for Meeting end event.
        /// </summary>
        /// <param name="meeting"></param>
        /// <param name="turnContext"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        protected override async Task OnTeamsMeetingEndAsync(MeetingEndEventDetails meeting, ITurnContext<IEventActivity> turnContext, CancellationToken cancellationToken)
        {
            var conversationStateAccessors = _conversationState.CreateProperty<MeetingData>(nameof(MeetingData));
            var conversationData = await conversationStateAccessors.GetAsync(turnContext, () => new MeetingData());
            await turnContext.SendActivityAsync(MessageFactory.Attachment(GetAdaptiveCardForMeetingEnd(meeting, conversationData,  turnContext, cancellationToken)));
            await SendToThreadAsync(_appId, _appSecret, turnContext.Activity.ServiceUrl, turnContext.Activity.Conversation.Id, "Meeting has ended");
        }

        /// <summary>
        /// Sample Adaptive card for Meeting Start event.
        /// </summary>
        private Bot.Schema.Attachment GetAdaptiveCardForMeetingStart(MeetingStartEventDetails meeting, 
            MeetingData conversationData,
             ITurnContext<IEventActivity> turnContext, 
             CancellationToken cancellationToken)
        {
            var tokenResponse = GetTokenResponse(turnContext, null, cancellationToken);

            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion("1.2"))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = meeting.Title  + "- started ",
                        Weight = AdaptiveTextWeight.Bolder,
                        Spacing = AdaptiveSpacing.Medium,
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = $"Start id: {meeting.Id} , User Token is {tokenResponse.Result.Token.ToString()}, conversation id: {turnContext.Activity.Conversation.Id}, service URL: {turnContext.Activity.ServiceUrl}",
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveOpenUrlAction
                    {
                        Title = "Join meeting",
                        Url = meeting.JoinUrl,
                    },
                },
            };

            return new Bot.Schema.Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }

        /// <summary>
        /// Sample Adaptive card for Meeting End event.
        /// </summary>
        private Bot.Schema.Attachment GetAdaptiveCardForMeetingEnd(MeetingEndEventDetails meeting, 
            MeetingData conversationData,
             ITurnContext<IEventActivity> turnContext, 
             CancellationToken cancellationToken)
        {

            TimeSpan meetingDuration = meeting.EndTime - conversationData.StartTime;
            var meetingDurationText = meetingDuration.Minutes < 1 ?
                  Convert.ToInt32(meetingDuration.Seconds) + "s"
                : Convert.ToInt32(meetingDuration.Minutes) + "min " + Convert.ToInt32(meetingDuration.Seconds) + "s";
            
            var tokenResponse = GetTokenResponse(turnContext, null, cancellationToken);
            
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion("1.2"))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = meeting.Title  + $"- ended",
                        Weight = AdaptiveTextWeight.Bolder,
                        Spacing = AdaptiveSpacing.Medium,
                    },
                     new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = $"End Time : {Convert.ToString(meeting.EndTime.ToLocalTime())}",
                                        Wrap = true,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = $"Total duration : {meetingDurationText}",
                                        Wrap = true,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = $"Meeting id: {meeting.Id} , User Token is {tokenResponse.Result.Token.ToString()}, conversation id: {turnContext.Activity.Conversation.Id}, service URL: {turnContext.Activity.ServiceUrl}",
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },
                }
            };

            return new Bot.Schema.Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }

        // Create the send policy for Microsoft Teams
        // For more information about these policies
        // see: http://www.thepollyproject.org/
        static IAsyncPolicy CreatePolicy() {
            // Policy for handling the short-term transient throttling.
            // Retry on throttling, up to 3 times with a 2,4,8 second delay between with a 0-1s jitter.
            var transientRetryPolicy = Policy
                    .Handle<ErrorResponseException>(ex => ex.Message.Contains("429"))
                    .WaitAndRetryAsync(
                        retryCount: 3, 
                        (attempt) => TimeSpan.FromSeconds(Math.Pow(2, attempt)) + TimeSpan.FromMilliseconds(_random.Next(0, 1000)));

            // Policy to avoid sending even more messages when the long-term throttling occurs.
            // After 5 messages fail to send, the circuit breaker trips & all subsequent calls will throw
            // a BrokenCircuitException for 10 minutes.
            // Note, in this application this cannot trip since it only sends one message at a time!
            // This is left in for completeness / demonstration purposes.
            var circuitBreakerPolicy = Policy
                .Handle<ErrorResponseException>(ex => ex.Message.Contains("429"))
                .CircuitBreakerAsync(exceptionsAllowedBeforeBreaking: 5, TimeSpan.FromMinutes(10));
            
            // Policy to wait and retry when long-term throttling occurs. 
            // This will retry a single message up to 5 times with a 10 minute delay between each attempt.
            // Note, in this application this cannot trip since the circuit breaker above cannot trip.
            // This is left in for completeness / demonstration purposes.
            var outerRetryPolicy = Policy
                .Handle<BrokenCircuitException>()
                .WaitAndRetryAsync(
                    retryCount: 5,
                    (_) => TimeSpan.FromMinutes(10));
            
            // Combine all three policies so that it will first attempt to retry short-term throttling (inner-most)
            // After 15 (5 messages, 3 failures each) consecutive failed attempts to send a message it will trip the circuit breaker
            // which will fail all messages for the next ten minutes. It will attempt to send messages up to 5 times for a total
            // wait of 50 minutes before failing a message.
            return
                outerRetryPolicy.WrapAsync(
                    circuitBreakerPolicy.WrapAsync(
                        transientRetryPolicy));
        }
        
        static readonly IAsyncPolicy RetryPolicy = CreatePolicy();
        
        static Task SendWithRetries(Func<Task> callback)
        {
            return RetryPolicy.ExecuteAsync(callback);
        }

         /// Send a message to a thread in a channel.
        public static async Task SendToThreadAsync(string appId, string appPassword, string serviceUrl, string conversationId, string message)
        {
            var activity = MessageFactory.Text(message);
            activity.Summary = message; // Ensure that the summary text is populated so the toast notifications aren't generic text.

            var credentials = new MicrosoftAppCredentials(appId, appPassword);

            var connectorClient = new ConnectorClient(new Uri(serviceUrl), credentials);
            await SendWithRetries(async () => 
                    await connectorClient.Conversations.SendToConversationAsync(conversationId, activity));
        }

    }
}

