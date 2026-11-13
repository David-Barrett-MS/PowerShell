#
# <# %SCRIPTNAME% #>
#
# By David Barrett, Microsoft Ltd. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

<#
.SYNOPSIS

.DESCRIPTION

.EXAMPLE

#>


param (
<# Auth.ps1 %PARAMS_START% #>
<# Auth.ps1 %PARAMS_END% #>

<# Logging.ps1 %PARAMS_START% #>
<# Logging.ps1 %PARAMS_END% #>

<# Graph.ps1 %PARAMS_START% #>
<# Graph.ps1 %PARAMS_END% #>

<# Mailbox.ps1 %PARAMS_START% #>
<# Mailbox.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$True,HelpMessage="Custom parameter defined by script.")]
    [ValidateNotNullOrEmpty()]
    [string]$CustomParameter = ""
)

$script:ScriptVersion = "1.0.0"

<# Logging.ps1 %FUNCTIONS_START% #>
<# Logging.ps1 %FUNCTIONS_END% #>

<# Auth.ps1 %FUNCTIONS_START% #>
<# Auth.ps1 %FUNCTIONS_END% #>

<# Graph.ps1 %FUNCTIONS_START% #>
<# Graph.ps1 %FUNCTIONS_END% #>

<# Mailbox.ps1 %FUNCTIONS_START% #>
<# Mailbox.ps1 %FUNCTIONS_END% #>


# Main code

Log "Completed"
