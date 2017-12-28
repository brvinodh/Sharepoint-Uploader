using GalaSoft.MvvmLight;
using SP.SpCommonFun;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharepointBulkUploadTool.ViewModel
{
    public class ViewModelItem<T> : ViewModelBase, IValidatableItem
    {
        private T propValue;

        private bool isTaskInProgress;

        private FontAwesome.WPF.FontAwesomeIcon icon;

        Func<T, Task> onValueChangeMethod;

        IDisplayWindowHandler messageHandler = null;

        private Predicate<T> validateFunction = null;

        public ViewModelItem(IDisplayWindowHandler messageHandler, Predicate<T> validateFunction = null)
        {
            this.messageHandler = messageHandler;

            this.validateFunction = validateFunction;
        }

        public ViewModelItem(Func<T, Task> onValueChangeCallBack, IDisplayWindowHandler messageHandler, Predicate<T> validateFunction = null) :this(messageHandler)
        {
            // this method would be invovked whenever the value has changed
            this.onValueChangeMethod = onValueChangeCallBack;
        }



        public T Value
        {
            get { return propValue; }
            set
            {
                this.SetValue(value);
            }
        }

        public bool IsInProgress
        {
            get { return isTaskInProgress; }
            set
            {
                isTaskInProgress = value;
                this.RaisePropertyChanged();
            }
        }

        public FontAwesome.WPF.FontAwesomeIcon ItemStatusIcon
        {
            get { return icon; }
            set { icon = value; this.RaisePropertyChanged(); }
        }

        public void SetItemStatus(FieldStatus status)
        {
            bool isInProgress = false;
            VMColor color = VMColor.Green;
            switch (status)
            {
                case FieldStatus.Success:
                    this.ItemStatusIcon = FontAwesome.WPF.FontAwesomeIcon.CheckCircle;
                    break;
                case FieldStatus.InProgress:
                    this.ItemStatusIcon = FontAwesome.WPF.FontAwesomeIcon.Spinner;
                    color = VMColor.Black;
                    isInProgress = true;
                    break;
                case FieldStatus.Failure:
                case FieldStatus.Error:
                case FieldStatus.FieldNullOrEmpty:
                    this.ItemStatusIcon = FontAwesome.WPF.FontAwesomeIcon.Exclamation;
                    color = VMColor.Red;
                    break;
            }

            this.IsInProgress = isInProgress;
            this.IconColor = color;
            this.RaisePropertyChanged(nameof(IconColor));
        }

        public VMColor IconColor { get; set; } = VMColor.NoColor;

        public bool IsValid { get; set; } = true;

        public bool Validate()
        {
            if (this.validateFunction != null)
            {
                this.IsValid = this.validateFunction(this.Value);                
            }
                    
            if(this.Value != null && this.IsValid)
            {
                this.SetItemStatus(FieldStatus.Success);
                return true;
            }

            // set the status 
            this.SetItemStatus(FieldStatus.FieldNullOrEmpty);
            return false;
        }

        private async void SetValue(T value)
        {
            try
            {
                propValue = value;

                if (this.onValueChangeMethod != null)
                {
                    this.SetItemStatus(FieldStatus.InProgress);
                    await this.onValueChangeMethod(value);

                    // set the status only if the item is still valid
                    if (this.IsValid)
                    {
                        this.SetItemStatus(FieldStatus.Success);
                    }                    
                }

                this.RaisePropertyChanged(nameof(this.Value));
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "An unexpected error occured, Please try restarting the application");
                this.SetItemStatus(FieldStatus.Error);
                this.messageHandler.ShowErrorMessage("We have encountered an error, Please refer to the below exception and contact the support team. \n" + ex.ToString(), "Unexpected error");
            }
           
        }

        public void SetValidity(bool isValid, string message = "")
        {
            this.IsValid = isValid;
            this.Validate();
        }
    }

    public enum FieldStatus
    {
        Success,

        Error,

        Failure,

        InProgress, 

        FieldNullOrEmpty
    }

    public enum VMColor
    {
        Red,

        Green,

        Black, 

        NoColor
    }

    public interface IValidatableItem
    {
        bool Validate();
    }

}
