using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Threading;
using System.Security.Cryptography;
using System.Security.Permissions;
using System.Security;
using System.Runtime.InteropServices;

namespace UIAUtomation
{
    public class UIAutomation
    {
        public AutomationElement GetRootElement()
        {
            return AutomationElement.RootElement;
        }

        public AutomationElement GetElement(AutomationElement rootElement, AutomationProperty property, object value)
        {
            return GetElement(rootElement, property, value, TreeScope.Children);
        }

        public AutomationElement GetElement(AutomationElement rootElement, AutomationProperty property, object value, TreeScope searchScope)
        {

            AutomationElement aeMainWindow = null;

            int numWaits = 0;
            do
            {
                aeMainWindow = rootElement.FindFirst(searchScope, new PropertyCondition(property, value));
                ++numWaits;
                Thread.Sleep(200);
            } while (aeMainWindow == null && numWaits < 50);
            return aeMainWindow;
        }

        public AutomationElement GetElementWithoutWait(AutomationElement rootElement, AutomationProperty property, object value)
        {
            return GetElementWithoutWait(rootElement, property, value, TreeScope.Children);
        }

        public AutomationElement GetElementWithoutWait(AutomationElement rootElement, AutomationProperty property, object value, TreeScope searchScope)
        {
            AutomationElement aeMainWindow = rootElement.FindFirst(searchScope, new PropertyCondition(property, value));
            return aeMainWindow;
        }

        public bool PressButtonOnWindow(AutomationElement element, AutomationProperty property, object value)
        {
            return PressButtonOnWindow(element, property, value, TreeScope.Children);
        }

        public bool PressButtonOnWindow(AutomationElement element, AutomationProperty property, object value, TreeScope treeScope)
        {
            try
            {
                //var window = GetElementWithoutWait(element, property, value,treeScope);
                //if (window == null)
                //    window = GetElement(element,property,value,treeScope);
                //if(window==null)
                //    return false;
                //else
                //{
                var buttonInvoke = GetInvokePattern(element, property, value, TreeScope.Descendants);
                buttonInvoke.Invoke();
                return true;
                // }
            }
            catch
            {
                return false;
            }
        }

        public InvokePattern GetInvokePattern(AutomationElement window, AutomationProperty property, object value, TreeScope treeScope)
        {
            AutomationElement aeButtonElement;
            int numWaits = 0;
            do
            {
                aeButtonElement = window.FindFirst(treeScope, this.GetPropertyCondition(property, value));
                ++numWaits;
                Thread.Sleep(500);
            } while (aeButtonElement == null && numWaits < 75);
            object objPattern;
            InvokePattern invokePatternObj;
            if (aeButtonElement.TryGetCurrentPattern(InvokePattern.Pattern, out objPattern))
            {
                invokePatternObj = objPattern as InvokePattern;
                return invokePatternObj;
            }
            return null;
        }

        public ValuePattern GetValuePatternWithoutWait(AutomationElement window, AutomationProperty property, object value, TreeScope searchScope)
        {
            AutomationElement aeTextBoxElement = window.FindFirst(searchScope, GetPropertyCondition(property, value));
            if (aeTextBoxElement == null)
                throw new ElementNotAvailableException("TextBoxElement Element not Available");
            object objPattern;
            ValuePattern valuePatternObj;
            if (aeTextBoxElement.TryGetCurrentPattern(ValuePattern.Pattern, out objPattern))
            {
                valuePatternObj = objPattern as ValuePattern;
                return valuePatternObj;
            }
            else
                throw new ElementNotEnabledException("The Value Pattern was not retrieved from the element");
        }

        public InvokePattern GetInvokePatternWithoutWait(AutomationElement window, AutomationProperty property, object value, TreeScope searchScope)
        {
            AutomationElement aeButtonElement = window.FindFirst(searchScope, GetPropertyCondition(property, value));
            if (aeButtonElement == null)
                throw new ElementNotAvailableException("Button Element not Available. Try the GetInvokePattern method which has a wait");
            object objPattern;
            InvokePattern invokePatternObj;
            if (aeButtonElement.TryGetCurrentPattern(InvokePattern.Pattern, out objPattern))
            {
                invokePatternObj = objPattern as InvokePattern;
                return invokePatternObj;
            }
            else
                throw new ElementNotEnabledException("The Invoke Pattern was not retrieved from the element");
        }

        public ExpandCollapsePattern GetExpandCollapsePattern(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope)
        {
            AutomationElement aeExpanderElement;
            int numWaits = 0;
            do
            {
                aeExpanderElement = GetFirstChildNode(element, property, value, searchScope);
                ++numWaits;
                Thread.Sleep(300);
            } while (aeExpanderElement == null && numWaits < 75);
            object objPattern;
            ExpandCollapsePattern togPattern;
            if (true == aeExpanderElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out objPattern))
            {
                togPattern = objPattern as ExpandCollapsePattern;
                return togPattern;
            }
            else
                return null;

        }

        public void SetValueOnTextBox(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope, string valueToBeSet)
        {
            var valuePattern = GetValuePatternWithoutWait(element, property, value, searchScope);
            if (valuePattern != null)
                valuePattern.SetValue(valueToBeSet);
        }

        public ExpandCollapsePattern GetExpandCollapsePatternWithoutWait(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope)
        {
            try
            {
                AutomationElement aeExpanderElement = GetFirstChildNode(element, property, value, searchScope);
                if (aeExpanderElement == null)
                    throw new ElementNotAvailableException("Expander Element not available. Try the GetExpandCollapsePattern which has a wait");
                object objPattern;
                ExpandCollapsePattern togPattern;
                if (true == element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out objPattern))
                {
                    togPattern = objPattern as ExpandCollapsePattern;
                    return togPattern;
                }
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public SelectionItemPattern GetSelectionItemPattern(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope)
        {
            AutomationElement aeSelectionPattern;
            int numWaits = 0;
            do
            {
                aeSelectionPattern = GetFirstChildNode(element, property, value, searchScope);
                ++numWaits;
                Thread.Sleep(300);
            } while (aeSelectionPattern == null && numWaits < 75);
            object objPattern;
            SelectionItemPattern selectionItemPattern;
            if (true == element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out objPattern))
            {
                selectionItemPattern = objPattern as SelectionItemPattern;
                return selectionItemPattern;
            }
            else
                return null;
        }

        public SelectionItemPattern GetSelectionItemWithoutWait(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope)
        {
            try
            {
                AutomationElement aeExpanderElement = GetFirstChildNode(element, property, value, searchScope);
                if (aeExpanderElement == null)
                    throw new ElementNotAvailableException("Expander Element not available. Try the GetSelectionItemPattern which has a wait");
                object objPattern;
                SelectionItemPattern togPattern;
                if (true == element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out objPattern))
                {
                    togPattern = objPattern as SelectionItemPattern;
                    return togPattern;
                }
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public TextPattern GetTextPatternWithoutWait(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope)
        {
            try
            {
                AutomationElement aeTextElement = GetFirstChildNode(element, property, value, searchScope);
                if (aeTextElement == null)
                    throw new ElementNotAvailableException("Text Element not available");
                object objPattern;
                TextPattern txtPattern;
                if (true == aeTextElement.TryGetCurrentPattern(TextPattern.Pattern, out objPattern))
                {
                    txtPattern = objPattern as TextPattern;
                    return txtPattern;
                }
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public string GetTextFromTextElement(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope)
        {
            var textPattern = GetTextPatternWithoutWait(element, property, value, searchScope);
            if (textPattern != null)
                return textPattern.DocumentRange.GetText(-1);
            return null;

        }

        public PropertyCondition GetPropertyCondition(AutomationProperty property, object value)
        {
            return new PropertyCondition(property, value);
        }

        public AutomationElementCollection GetAllChildNodes(AutomationElement element, AutomationProperty automationProperty, object value, TreeScope treeScope)
        {
            var allChildNodes = element.FindAll(treeScope, GetPropertyCondition(automationProperty, value));
            if (allChildNodes == null)
                throw new ElementNotAvailableException("Not able to find the child nodes of the element");
            return allChildNodes;
        }

        public AutomationElement GetFirstChildNode(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope)
        {
            var firstchildNode = element.FindFirst(searchScope, GetPropertyCondition(property, value));
            if (firstchildNode == null)
                throw new ElementNotAvailableException("Not able to find the first child node of the element");
            return firstchildNode;
        }

        public bool FirstChildTextNodeContains(AutomationElement element, string toCheck)
        {
            var firstTextNode = this.GetFirstChildNode(element, AutomationElement.ControlTypeProperty, ControlType.Text, TreeScope.Children);
            if (firstTextNode.Current.Name == toCheck)
                return true;
            else
                return false;
        }

        public void SelectComboBoxItem(AutomationElement element, int indexToBeSelected, AutomationProperty property, object value, TreeScope searchScope)
        {
            var allChildren = this.ExpandComboBoxViewAndReturnChildren(element, property, value, searchScope);
            if (allChildren.Count < indexToBeSelected)
                throw new Exception("The combobox has fewer items than those that need to be selected");
            var itemToBeSelected = allChildren[indexToBeSelected];
            var selectPattern = this.GetSelectionItemPattern(itemToBeSelected, AutomationElement.ControlTypeProperty, ControlType.ListItem, TreeScope.Element);
            if (selectPattern != null)
                selectPattern.Select();
            var togglePattern = this.GetExpandCollapsePattern(element, property, value, TreeScope.Descendants);
            if (togglePattern != null)
                togglePattern.Collapse();
        }

        public AutomationElementCollection ExpandComboBoxViewAndReturnChildren(AutomationElement element, AutomationProperty property, object value, TreeScope searchScope)
        {
            try
            {
                var comboBox = GetFirstChildNode(element, AutomationElement.AutomationIdProperty, value, searchScope);
                var expandPattern = GetExpandCollapsePattern(comboBox, AutomationElement.ControlTypeProperty, ControlType.ComboBox, TreeScope.Element);
                if (expandPattern == null)
                    throw new ElementNotAvailableException("Couldnt Find Expand Pattern in combobox");
                expandPattern.Expand();
                return this.GetAllChildNodes(comboBox, AutomationElement.ControlTypeProperty, ControlType.ListItem, TreeScope.Children);
            }
            catch (Exception e1)
            {
                throw e1;
            }
        }

        public void WorkaroundPopulateControlTree(int x1, int y1)
        {
            uint x = (uint)x1;
            uint y = (uint)y1;
            NativeMethods.SetCursorPos(x1, y1);
            Thread.Sleep(200);
            NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTDOWN | NativeMethods.MOUSEEVENTF_LEFTUP, x, y, 0, 0);
        }

        public bool WaitTillElementSelected(AutomationElement aeMainWindow, AutomationProperty automationProperty, object value, TreeScope treeScope)
        {
            var btnStartAnalysis = GetElement(aeMainWindow, automationProperty, value, treeScope);
            int numWait = 0;
            do
            {
                if (btnStartAnalysis.Current.IsEnabled == true)
                    return true;
                Thread.Sleep(2000);
            } while (btnStartAnalysis.Current.IsEnabled == false && numWait < 90);

            return false;
        }

    }
}
