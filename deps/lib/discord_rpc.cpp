#include "../include/discord_rpc.h"
#include <cstdio>
#include <windows.h>
#include <cstring>

static DiscordEventHandlers g_handlers = {};
static bool g_initialized = false;

void Discord_Initialize(const char* applicationId, DiscordEventHandlers* handlers, int autoRegister, const char* optionalSteamId) {
    if (handlers) {
        g_handlers = *handlers;
    }
    g_initialized = true;

    if (g_handlers.ready) {
        // Dapatkan username dari environment atau Windows
        char username[256] = "Unknown";
        DWORD size = sizeof(username);
        
        if (GetEnvironmentVariableA("USERNAME", username, size) == 0) {
            strcpy_s(username, sizeof(username), "Discord User");
        }

        DiscordUser user = {
            username,
            "0001",
            "123456789",
            nullptr
        };
        g_handlers.ready(&user);
    }
}

void Discord_Shutdown(void) {
    g_initialized = false;
}

void Discord_RunCallbacks(void) {
    // Callbacks processing (non-blocking)
}

void Discord_UpdatePresence(const DiscordRichPresence* presence) {
    if (!g_initialized || !presence) return;

    // Log for debugging
    if (presence->details) {
        printf("Discord RPC Details: %s\n", presence->details);
    }
    if (presence->state) {
        printf("Discord RPC State: %s\n", presence->state);
    }
}

void Discord_ClearPresence(void) {
    if (!g_initialized) return;
}

void Discord_Respond(const char* userId, int reply) {
    if (!g_initialized || !userId) return;
}