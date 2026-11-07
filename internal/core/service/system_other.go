//go:build !windows
// +build !windows

package service

import (
	"time"
)

type unavailableSCM struct{}

func System() Deps {
	return Deps{SCM: unavailableSCM{}, Netsh: Netsh, Dispatch: func(string, func(Session)) error { return ErrNotService }, Sleep: time.Sleep, Now: time.Now}
}
func (unavailableSCM) IsAdmin() bool                   { return false }
func (unavailableSCM) Exists(string) (bool, error)     { return false, unavailableError("Exists") }
func (unavailableSCM) Create(string, Settings) error   { return unavailableError("Create") }
func (unavailableSCM) Update(string, Settings) error   { return unavailableError("Update") }
func (unavailableSCM) Config(string) (Settings, error) { return Settings{}, unavailableError("Config") }
func (unavailableSCM) Start(string) error              { return unavailableError("Start") }
func (unavailableSCM) Stop(string) error               { return unavailableError("Stop") }
func (unavailableSCM) Delete(string) error             { return unavailableError("Delete") }
func (unavailableSCM) State(string) (State, error)     { return 0, unavailableError("State") }

func unavailableError(op string) error {
	return &Error{Code: 1, Msg: op + " is only available on Windows"}
}
