from pennsieve2.pennsieve import Pennsieve
from .exceptions import PennsieveAgentError, PennsieveAgentRPCError, PennsieveAccountInvalid

def connect_pennsieve_client(account_name):
    """
        Connects to Pennsieve Python client to the Agent and returns the initialized Pennsieve object.
    """
    try:
        return Pennsieve(profile_name=account_name)
    except Exception as e:
        if "End of TCP stream" in str(e):
            raise PennsieveAgentRPCError(f"Could not connect to the Pennsieve Agent error coming from the Agent. This can likely be resolved by retrying. If the issue persists, please contact the SODA team and they will reach out to Pennsieve to help resolve this.") from e
        elif isinstance(e, AttributeError):
            raise Exception("The Pennsieve Agent cannot access datasets but needs to in order to work. Please try again. If the issue persists, please contact the SODA team. The SODA team will contact Pennsieve to help resolve this issue.") from e
        elif "Profile not found" in str(e):
            raise PennsieveAccountInvalid(account_name) from e
        elif "Error connecting to server" in str(e):
            raise PennsieveAgentError(f"Could not connect to the Pennsieve agent: {e}") from e
        else:
            raise PennsieveAgentError("An unknown error occurred when trying to connect to the Pennsieve Agent. If this issue persists, please contact the SODA team.") from e