using System.Drawing;
using System.Runtime.InteropServices;
using NUnit.Framework;
using OfficeGbAutomation.Common;

namespace Word.Tests;

[TestFixture, Category("Infrastructure")]
public sealed class StartupTests
{
    [TestCase("NetUInetpane", "OfficePLUS", true)]
    [TestCase("NetUInetpane", "Copilot", true)]
    [TestCase("NetUInetpane", "officeplus", true)]
    [TestCase("NetUInetpane", "Find and Replace", false)]
    [TestCase("NetUInetpane", "OfficePLUS templates", false)]
    [TestCase("Button", "OfficePLUS", false)]
    [TestCase("PPTFrameClass", "Copilot", false)]
    public void Only_exact_known_task_pane_markers_are_eligible(string className, string name, bool expected) =>
        Assert.That(OfficeStartup.IsKnownTaskPane(className, name), Is.EqualTo(expected));

    [TestCase("NetUINativeHWNDHost", "Copilot", "Copilot", true)]
    [TestCase("NetUINativeHWNDHost", "OfficePLUS", "OfficePLUS", true)]
    [TestCase("NetUINativeHWNDHost", "Find and Replace", "Copilot", false)]
    [TestCase("NetUINativeHWNDHost", "Copilot", "OfficePLUS", false)]
    [TestCase("PPTFrameClass", "Copilot", "Copilot", false)]
    public void Close_header_must_match_the_identified_task_pane(string className, string name, string paneName, bool expected) =>
        Assert.That(OfficeStartup.IsPaneHeader(className, name, paneName), Is.EqualTo(expected));

    [TestCase("Close pane", true)]
    [TestCase("\u5173\u95ed\u7a97\u683c", true)]
    [TestCase("Close", false)]
    [TestCase("Restart", false)]
    public void Only_pane_close_buttons_are_eligible(string name, bool expected) =>
        Assert.That(OfficeStartup.IsPaneCloseButton(name), Is.EqualTo(expected));

    [TestCase("Copilot", "sessionPaneCloseButton", true)]
    [TestCase("copilot", "sessionPaneCloseButton", true)]
    [TestCase("OfficePLUS", "sessionPaneCloseButton", false)]
    [TestCase("Find and Replace", "sessionPaneCloseButton", false)]
    [TestCase("Copilot", "Close", false)]
    [TestCase("Copilot", "", false)]
    public void Modern_chat_close_requires_Copilot_and_exact_control_id(
        string paneName, string automationId, bool expected) =>
        Assert.That(OfficeStartup.IsCopilotChatCloseButton(paneName, automationId), Is.EqualTo(expected));

    private static void Wait(Func<bool> condition, string description) =>
        OfficeStartup.WaitFor(condition, description, TimeSpan.FromMilliseconds(100), _ => { });

    [TestCase(null, null)]
    [TestCase(true, null)]
    [TestCase(true, true)]
    [TestCase(null, false)]
    [TestCase(false, false)]
    public void Hidden_or_unknown_pane_visibility_never_invokes_Close(bool? offscreen, bool? nativeVisible)
    {
        var closes = 0;
        var messages = new List<string>();
        OfficeStartup.CloseOptionalPane("OfficePLUS",
            () => OfficeStartup.IsVisiblyOpen(offscreen, nativeVisible,
                new Rectangle(100, 100, 200, 200), new Rectangle(0, 0, 1000, 800)),
            () => { closes++; return true; }, Wait, messages.Add);
        Assert.That(closes, Is.Zero);
        Assert.That(messages, Is.Empty);
    }

    [TestCase(false, null)]
    [TestCase(false, true)]
    [TestCase(null, true)]
    public void Pane_visibility_requires_positive_UIA_or_verified_native_evidence(bool? offscreen, bool? nativeVisible) =>
        Assert.That(OfficeStartup.IsVisiblyOpen(offscreen, nativeVisible,
            new Rectangle(100, 100, 200, 200), new Rectangle(0, 0, 1000, 800)), Is.True);

    [TestCase(0, 0, 0, 0)]
    [TestCase(100, 100, 0, 200)]
    [TestCase(100, 100, 200, 0)]
    [TestCase(1500, 100, 200, 200)]
    [TestCase(-300, 100, 200, 200)]
    public void Empty_or_outside_pane_bounds_are_not_visible(int x, int y, int width, int height) =>
        Assert.That(OfficeStartup.IsVisiblyOpen(false, true,
            new Rectangle(x, y, width, height), new Rectangle(0, 0, 1000, 800)), Is.False);

    [Test]
    public void Already_closed_or_absent_pane_never_invokes_Close()
    {
        var closes = 0;
        var messages = new List<string>();
        OfficeStartup.CloseOptionalPane("OfficePLUS", () => false, () => { closes++; return true; },
            Wait, messages.Add);
        Assert.That(closes, Is.Zero);
        Assert.That(messages, Is.Empty);
    }

    [Test]
    public void Open_pane_is_closed_and_its_disappearance_verified()
    {
        var open = true;
        var closes = 0;
        var messages = new List<string>();
        OfficeStartup.CloseOptionalPane("OfficePLUS", () => open, () =>
        {
            closes++;
            open = false;
            return true;
        }, Wait, messages.Add);
        Assert.That(closes, Is.EqualTo(1));
        Assert.That(open, Is.False);
        Assert.That(messages.Single(), Does.Contain("no add-in settings are changed"));
    }

    [Test]
    public void Pane_disappearing_before_Close_is_not_clicked()
    {
        var probes = 0;
        var closes = 0;
        OfficeStartup.CloseOptionalPane("Copilot", () => ++probes == 1,
            () => { closes++; return true; }, Wait, _ => { });
        Assert.That(closes, Is.Zero);
    }

    [Test]
    public void Missing_safe_Close_control_fails_with_actionable_message()
    {
        var error = Assert.Throws<TimeoutException>(() =>
            OfficeStartup.CloseOptionalPane("OfficePLUS", () => true, () => false, Wait, _ => { }));
        Assert.That(error!.Message, Does.Contain("OfficePLUS").And.Contain("close the pane manually"));
    }

    [Test]
    public void Close_that_does_not_dismiss_the_pane_fails_without_repeated_clicks()
    {
        var closes = 0;
        var error = Assert.Throws<TimeoutException>(() =>
            OfficeStartup.CloseOptionalPane("Copilot", () => true, () => { closes++; return true; },
                Wait, _ => { }));
        Assert.That(closes, Is.EqualTo(1));
        Assert.That(error!.Message, Does.Contain("Copilot task pane to disappear"));
    }

    [TestCase(unchecked((int)0x80040201))]
    [TestCase(unchecked((int)0x80010001))]
    [TestCase(unchecked((int)0x8001010A))]
    public void Transient_UIA_failures_are_logged_and_retried(int hresult)
    {
        var calls = 0;
        var messages = new List<string>();
        OfficeStartup.WaitFor(() =>
        {
            if (++calls < 3)
                throw new COMException("Transient test failure", hresult);
            return true;
        }, "startup", TimeSpan.FromSeconds(3), messages.Add);
        Assert.That(calls, Is.EqualTo(3));
        Assert.That(messages, Has.Count.EqualTo(1));
        Assert.That(messages[0], Does.Contain($"0x{hresult:X8}"));
    }

    [Test]
    public void Persistent_transient_failure_times_out_with_original_error()
    {
        var original = new COMException("Persistent test failure", unchecked((int)0x80040201));
        var error = Assert.Throws<TimeoutException>(() => OfficeStartup.WaitFor(() => throw original,
            "startup", TimeSpan.FromMilliseconds(100), _ => { }));
        Assert.That(error!.InnerException, Is.SameAs(original));
        Assert.That(error.Message, Does.Contain("0x80040201"));
    }

    [Test]
    public void Unrelated_COM_failures_are_not_retried()
    {
        var calls = 0;
        var original = new COMException("Access denied", unchecked((int)0x80070005));
        var error = Assert.Throws<COMException>(() => OfficeStartup.WaitFor(() =>
        {
            calls++;
            throw original;
        }, "startup", TimeSpan.FromSeconds(1), _ => { }));
        Assert.That(error, Is.SameAs(original));
        Assert.That(calls, Is.EqualTo(1));
    }
}
