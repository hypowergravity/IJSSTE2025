// Registration form functionality
document.addEventListener('DOMContentLoaded', function() {
    const registrationForm = document.getElementById('registrationForm');
    
    // Form validation
    function validateForm() {
        const requiredFields = registrationForm.querySelectorAll('input[required], select[required]');
        let isValid = true;
        
        requiredFields.forEach(field => {
            if (!field.value.trim()) {
                field.style.borderColor = '#dc2626';
                isValid = false;
            } else {
                field.style.borderColor = '#d1d5db';
            }
        });
        
        // Email validation
        const emailField = document.getElementById('email');
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (emailField.value && !emailRegex.test(emailField.value)) {
            emailField.style.borderColor = '#dc2626';
            isValid = false;
        }
        
        // Terms checkbox validation
        const termsCheckbox = document.getElementById('terms');
        if (!termsCheckbox.checked) {
            termsCheckbox.parentElement.style.color = '#dc2626';
            isValid = false;
        } else {
            termsCheckbox.parentElement.style.color = '';
        }
        
        return isValid;
    }
    
    // Show success message
    function showSuccessMessage() {
        const successMessage = document.createElement('div');
        successMessage.className = 'success-message';
        successMessage.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: #059669;
            color: white;
            padding: 1rem 2rem;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            z-index: 1000;
            animation: slideIn 0.3s ease-out;
        `;
        successMessage.innerHTML = `
            <strong>Registration Submitted!</strong><br>
            You will receive a confirmation email shortly.
        `;
        
        document.body.appendChild(successMessage);
        
        setTimeout(() => {
            successMessage.remove();
        }, 5000);
    }
    
    if (registrationForm) {
        registrationForm.addEventListener('submit', function(e) {
            e.preventDefault();
            
            if (validateForm()) {
                // Simulate form submission
                const submitButton = registrationForm.querySelector('button[type="submit"]');
                const originalText = submitButton.textContent;
                
                submitButton.disabled = true;
                submitButton.textContent = 'Processing...';
                
                setTimeout(() => {
                    showSuccessMessage();
                    submitButton.disabled = false;
                    submitButton.textContent = originalText;
                    
                    // Reset form
                    registrationForm.reset();
                }, 2000);
            } else {
                // Show error message
                const errorMessage = document.createElement('div');
                errorMessage.style.cssText = `
                    background: #fef2f2;
                    border: 1px solid #fecaca;
                    color: #dc2626;
                    padding: 1rem;
                    border-radius: 6px;
                    margin-bottom: 1rem;
                `;
                errorMessage.textContent = 'Please fill in all required fields correctly.';
                
                const existingError = registrationForm.querySelector('.error-message');
                if (existingError) {
                    existingError.remove();
                }
                
                errorMessage.className = 'error-message';
                registrationForm.insertBefore(errorMessage, registrationForm.firstChild);
                
                setTimeout(() => {
                    errorMessage.remove();
                }, 5000);
            }
        });
    }
    
    // Add CSS animations
    const style = document.createElement('style');
    style.textContent = `
        @keyframes slideIn {
            from {
                transform: translateX(100%);
                opacity: 0;
            }
            to {
                transform: translateX(0);
                opacity: 1;
            }
        }
        
        .form-group input:focus,
        .form-group select:focus,
        .form-group textarea:focus {
            transform: translateY(-2px);
            transition: all 0.3s ease;
        }
        
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
        }
        
        .hotel-card:hover,
        .facility-card:hover,
        .category-card:hover {
            box-shadow: 0 8px 25px rgba(0, 0, 0, 0.15);
        }
    `;
    document.head.appendChild(style);
});