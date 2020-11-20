#!
# -*- coding: utf-8 -*-

#from __futur__ import absolute_import

import uno

from .oauth2lib import InteractionRequest
from .unotools import getInteractionHandler


def getUserNameFromHandler(ctx, url, source, message=''):
    username = ''
    handler = getInteractionHandler(ctx)
    interaction = InteractionRequest(source, url, message)
    if handler.handleInteractionRequest(interaction):
        continuation = interaction.getContinuations()[-1]
        username = continuation.getUserName()
    return username

def getOAuth2UserName(ctx, source, url, message=''):
    username = ''
    handler = getInteractionHandler(ctx)
    interaction = InteractionRequest(source, url, '', '', message)
    if handler.handleInteractionRequest(interaction):
        continuation = interaction.getContinuations()[-1]
        username = continuation.getUserName()
    return username

def getOAuth2Token(ctx, source, url, user, format=''):
    token = ''
    handler = getInteractionHandler(ctx)
    interaction = InteractionRequest(source, url, user, format, '')
    if handler.handleInteractionRequest(interaction):
        continuation = interaction.getContinuations()[-1]
        token = continuation.getToken()
    return token
