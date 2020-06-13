import { OnlineMeetingInput, OnlineMeeting } from './models';
//import { msalApp } from '../auth/msalApp';
import axios from 'axios';
import * as moment from 'moment';

import * as debug from "debug";
const log = debug("msteams");

export function createMeetingService() {
    return {
        async createMeeting(meeting: OnlineMeetingInput, accessToken: string) {
            //   let token;
            //   try {
            //     token = await msalApp.acquireTokenSilent({
            //       scopes: ['OnlineMeetings.ReadWrite']
            //     });
            //   } catch (ex) {
            //     token = await msalApp.acquireTokenPopup({
            //       scopes: ['OnlineMeetings.ReadWrite']
            //     });
            //   }

            const requestBody = {
                startDateTime: meeting.startDateTime?.toISOString(),
                endDateTime: meeting.endDateTime?.toISOString(),
                subject: meeting.subject
            };
            log("token: " + accessToken);
            log("Graph call: Me");
            const resMe = await axios.get('https://graph.microsoft.com/v1.0/me',
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        'Content-type': 'application/json'
                    }
                });
            log("Response: " + JSON.stringify(resMe.data));

            log("Graph call: onlineMeetings");
            const response = await axios.post(
                'https://graph.microsoft.com/v1.0/me/onlineMeetings',
                requestBody,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        'Content-type': 'application/json'
                    }
                }
            );
            log("Response: " + JSON.stringify(response.data));

            const preview = decodeURIComponent(
                (response.data.joinInformation.content?.split(',')?.[1] ?? '').replace(
                    /\+/g,
                    '%20'
                )
            );

            log("Graph call: Create Outlook Event");
            const eventBody = {
                subject: "Prep for customer meeting",
                body: {
                    contentType: "HTML",
                    content: "Does this time work for you?"
                },
                start: {
                    dateTime: "2020-06-14T13:00:00",
                    timeZone: "Pacific Standard Time"
                },
                end: {
                    dateTime: "2020-06-14T14:00:00",
                    timeZone: "Pacific Standard Time"
                },
                location: {
                    displayName: "Naga Meeting for Sample"
                },
                attendees: [
                    {
                        emailAddress: {
                            address: "nagarajan_s05@msnextlife.OnMicrosoft.com",
                            name: "Nagarajan Subramani"
                        },
                        type: "required"
                    }
                ],
                allowNewTimeProposals: true,
                isOnlineMeeting: true,
                onlineMeetingProvider: "teamsForBusiness"
            };
            const resEvent = await axios.post(
                'https://graph.microsoft.com/v1.0/me/events',
                eventBody,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        'Content-type': 'application/json'
                    }
                }
            );
            log("Response: " + JSON.stringify(resEvent.data));


            const createdMeeting = {
                id: response.data.id,
                creationDateTime: moment(response.data.creationDateTime),
                subject: response.data.subject,
                joinUrl: response.data.joinUrl,
                joinWebUrl: response.data.joinWebUrl,
                startDateTime: moment(response.data.startDateTime),
                endDateTime: moment(response.data.endDateTime),
                conferenceId: response.data.audioConferencing?.conferenceId || '',
                tollNumber: response.data.audioConferencing?.tollNumber || '',
                tollFreeNumber: response.data.audioConferencing?.tollFreeNumber || '',
                dialinUrl: response.data.audioConferencing?.dialinUrl || '',
                videoTeleconferenceId: response.data.videoTeleconferenceId,
                preview
            } as OnlineMeeting;

            return createdMeeting;
        }
    };
}
