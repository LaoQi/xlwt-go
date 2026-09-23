package xlwt

import "errors"

// errorsIs is a tiny wrapper so tests read the same regardless of the errors
// package being imported in each file.
func errorsIs(err, target error) bool { return errors.Is(err, target) }
